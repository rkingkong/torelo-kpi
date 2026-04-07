#!/usr/bin/env python3
"""
AI-Powered Email Search Tool for Torelo - FIXED VERSION
Corrects 400 errors and consolidates configuration
"""

import os
import json
import asyncio
import base64
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional, Set
import re
from pathlib import Path
import logging
from concurrent.futures import ThreadPoolExecutor
import aiohttp
import requests
from msal import ConfidentialClientApplication
import anthropic
from flask import Flask, request, jsonify, send_file, Response
from flask_cors import CORS
import configparser
import time

# ============================================================================
# CONSOLIDATED CONFIGURATION
# ============================================================================

# Load config.ini for credentials only
config = configparser.ConfigParser()
config.read('config.ini')

# Azure AD / Microsoft Graph settings (from config.ini)
CLIENT_ID = config['Azure']['client_id']
CLIENT_SECRET = config['Azure']['client_secret'] 
TENANT_ID = config['Azure']['tenant_id']
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

# Claude API settings (from config.ini)
CLAUDE_API_KEY = config['Claude']['api_key']

# ============================================================================
# SEARCH CONFIGURATION - ALL IN ONE PLACE
# ============================================================================

# Search Speed Mode
SEARCH_MODE = "NORMAL"  # Options: "FAST", "NORMAL", "DEEP"

# Search Configuration by Mode
SEARCH_CONFIG = {
    "FAST": {
        "max_search_terms": 3,      # Use only top 3 most relevant terms
        "max_users_to_search": 5,   # Search only 5 priority users
        "days_to_look_back": 30,    # Look back 30 days
        "emails_per_term": 10,      # Get 10 emails per search term
        "concurrent_searches": 3,    # Run 3 searches at once (avoid rate limits)
        "skip_attachments": True,   # Don't fetch attachment details
        "simple_consensus": True    # Use simple consensus building
    },
    "NORMAL": {
        "max_search_terms": 5,      # Use top 5 relevant terms
        "max_users_to_search": 15,  # Search 15 users
        "days_to_look_back": 90,    # Look back 90 days
        "emails_per_term": 15,      # Get 15 emails per search term
        "concurrent_searches": 3,    # Run 3 searches at once
        "skip_attachments": False,  # Fetch attachment details
        "simple_consensus": False   # Use AI consensus
    },
    "DEEP": {
        "max_search_terms": 10,     # Use more search terms
        "max_users_to_search": 50,  # Search many users
        "days_to_look_back": 365,   # Look back full year
        "emails_per_term": 25,      # Get more emails per term
        "concurrent_searches": 2,    # Run 2 searches (to avoid rate limits)
        "skip_attachments": False,  # Fetch all attachments
        "simple_consensus": False   # Full consensus building
    }
}

# Get current configuration
CURRENT_CONFIG = SEARCH_CONFIG[SEARCH_MODE]

# Priority Users - These are checked first and guaranteed to be searched
PRIORITY_USERS = [
    "admin@torelo.net",
    "rkong@torelo.net",  # From your config
    "jmkong@torelo.net",
    "arq@torelo.net",
    "sgarcia@torelo.net",
    "alopez@torelo.net"
]

# Spanish Terms Mapping (from config.ini + more)
SPANISH_TO_ENGLISH = {
    'proyecto': 'project',
    'inventario': 'inventory',
    'reporte': 'report',
    'reunion': 'meeting',
    'reunión': 'meeting',
    'presupuesto': 'budget',
    'cliente': 'client',
    'factura': 'invoice',
    'pedido': 'order',
    'entrega': 'delivery',
    'pago': 'payment',
    'tarea': 'task',
    'fecha': 'date',
    'archivo': 'file',
    'documento': 'document',
    'urgente': 'urgent',
    'planos': 'plans',
    'hidráulico': 'hydraulic',
    'cronograma': 'schedule',
    'actualizado': 'updated',
    'último': 'latest',
    'cimentación': 'foundation',
    'torre': 'tower',
    'aprobación': 'approval',
    'pendiente': 'pending',
    'estado': 'status'
}

# Spanish-Only Mode
SPANISH_PRIORITY = True

# Cache Configuration
ENABLE_CACHE = True
CACHE_TTL_SECONDS = 300  # 5 minutes

# ============================================================================
# END OF CONFIGURATION
# ============================================================================

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Initialize Flask app
app = Flask(__name__)
CORS(app, origins=['*'])

# Global variables
token_cache = {
    'token': None,
    'expiry': None
}
query_cache = {} if ENABLE_CACHE else None


class OptimizedEmailSearchService:
    """Optimized email search service with correct API syntax"""
    
    def __init__(self):
        self.msal_app = ConfidentialClientApplication(
            CLIENT_ID,
            authority=AUTHORITY,
            client_credential=CLIENT_SECRET
        )
        self.claude_client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)
        self.results_dir = Path('search_results')
        self.results_dir.mkdir(exist_ok=True)
        
    def get_access_token(self):
        """Get or refresh access token with caching"""
        global token_cache
        
        if token_cache['token'] and token_cache['expiry']:
            if datetime.now() < token_cache['expiry']:
                return token_cache['token']
        
        result = self.msal_app.acquire_token_for_client(scopes=SCOPE)
        
        if "access_token" in result:
            token_cache['token'] = result['access_token']
            token_cache['expiry'] = datetime.now() + timedelta(minutes=55)
            return result['access_token']
        else:
            raise Exception(f"Failed to get access token: {result.get('error_description', 'Unknown error')}")
    
    def get_smart_search_terms(self, query: str) -> List[str]:
        """Get smart, limited search terms based on query"""
        query_lower = query.lower()
        terms = []
        
        # Extract words from query
        words = query_lower.split()
        
        # Skip common Spanish/English words
        skip_words = {'de', 'la', 'el', 'los', 'las', 'del', 'para', 'con', 'que', 'en', 'y', 'a', 
                     'the', 'of', 'and', 'or', 'for', 'in', 'to', 'is', 'are', 'was', 'were'}
        
        # Process each word
        for word in words:
            if word not in skip_words and len(word) > 2:
                terms.append(word)
                
                # Also add English translation if Spanish
                if word in SPANISH_TO_ENGLISH:
                    english_term = SPANISH_TO_ENGLISH[word]
                    if english_term not in terms:
                        terms.append(english_term)
        
        # Common query patterns - add specific terms
        if 'hidráulico' in query_lower or 'hydraulic' in query_lower:
            if 'water' not in terms: terms.append('water')
            if 'plumbing' not in terms: terms.append('plumbing')
        
        if 'torre 3' in query_lower:
            if 'tower 3' not in terms: terms.append('tower 3')
            if 'project' not in terms: terms.append('project')
        
        if 'cronograma' in query_lower:
            if 'schedule' not in terms: terms.append('schedule')
            if 'timeline' not in terms: terms.append('timeline')
        
        # Limit terms
        return terms[:CURRENT_CONFIG['max_search_terms']]
    
    async def get_active_users(self) -> List[str]:
        """Get users with active mailboxes"""
        # Start with priority users
        active_users = PRIORITY_USERS.copy()
        
        if len(active_users) >= CURRENT_CONFIG['max_users_to_search']:
            return active_users[:CURRENT_CONFIG['max_users_to_search']]
        
        # Get more users from Graph API if needed
        headers = {
            'Authorization': f'Bearer {self.get_access_token()}',
            'Content-Type': 'application/json'
        }
        
        try:
            # Get users with mailboxes enabled
            url = f"{GRAPH_API_BASE}/users?$select=userPrincipalName,mail,mailEnabled&$filter=mailEnabled eq true&$top=50"
            
            async with aiohttp.ClientSession() as session:
                async with session.get(url, headers=headers, timeout=aiohttp.ClientTimeout(total=10)) as resp:
                    if resp.status == 200:
                        data = await resp.json()
                        
                        for user in data.get('value', []):
                            email = user.get('userPrincipalName') or user.get('mail')
                            if email and '@' in email and email not in active_users:
                                # Check if it's a real user email (not service account)
                                if not any(skip in email.lower() for skip in ['noreply', 'service', 'system']):
                                    active_users.append(email)
                            
                            if len(active_users) >= CURRENT_CONFIG['max_users_to_search']:
                                break
                        
                        return active_users[:CURRENT_CONFIG['max_users_to_search']]
                    else:
                        logger.warning(f"Failed to get users: {resp.status}")
                        
        except Exception as e:
            logger.error(f"Error getting users: {str(e)}")
        
        # Return what we have
        return active_users[:CURRENT_CONFIG['max_users_to_search']]
    
    async def search_single_user_term(self, session, user_email: str, term: str, headers: dict) -> List[Dict]:
        """Search a single term for a single user with correct syntax"""
        try:
            url = f"{GRAPH_API_BASE}/users/{user_email}/messages"
            
            # Use correct search syntax - just the term, no date filter in $search
            params = {
                '$search': f'"{term}"',
                '$select': 'id,subject,bodyPreview,from,toRecipients,receivedDateTime,hasAttachments',
                '$top': CURRENT_CONFIG['emails_per_term'],
                '$orderby': 'receivedDateTime desc'  # Sort by date
            }
            
            # Add date filter using $filter instead of $search
            if CURRENT_CONFIG['days_to_look_back'] < 365:
                date_filter = (datetime.now() - timedelta(days=CURRENT_CONFIG['days_to_look_back'])).strftime('%Y-%m-%dT00:00:00Z')
                params['$filter'] = f"receivedDateTime ge {date_filter}"
            
            if not CURRENT_CONFIG['skip_attachments']:
                params['$expand'] = 'attachments($select=id,name,size)'
            
            async with session.get(url, headers=headers, params=params, timeout=aiohttp.ClientTimeout(total=10)) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    emails = data.get('value', [])
                    
                    # Add metadata to each email
                    for email in emails:
                        email['searchedUser'] = user_email
                        email['matchedTerm'] = term
                    
                    logger.debug(f"Found {len(emails)} emails for {user_email} with term '{term}'")
                    return emails
                    
                elif resp.status == 404:
                    logger.debug(f"User {user_email} not found or no mailbox")
                    return []
                    
                elif resp.status == 400:
                    error_text = await resp.text()
                    logger.error(f"Bad request for {user_email}: {error_text}")
                    return []
                    
                else:
                    logger.warning(f"Error {resp.status} searching {user_email} for '{term}'")
                    return []
                    
        except asyncio.TimeoutError:
            logger.warning(f"Timeout searching {user_email} for '{term}'")
            return []
        except Exception as e:
            logger.error(f"Exception searching {user_email} for '{term}': {str(e)}")
            return []
    
    async def concurrent_search(self, users: List[str], terms: List[str]) -> List[Dict]:
        """Perform concurrent searches across users and terms"""
        headers = {
            'Authorization': f'Bearer {self.get_access_token()}',
            'Content-Type': 'application/json'
        }
        
        all_results = []
        
        async with aiohttp.ClientSession() as session:
            # Create all search tasks
            tasks = []
            for user in users:
                for term in terms:
                    task = self.search_single_user_term(session, user, term, headers)
                    tasks.append(task)
            
            # Run searches concurrently with limit
            semaphore = asyncio.Semaphore(CURRENT_CONFIG['concurrent_searches'])
            
            async def limited_search(task):
                async with semaphore:
                    return await task
            
            # Execute all searches
            results = await asyncio.gather(*[limited_search(task) for task in tasks], return_exceptions=True)
            
            # Collect results
            for result in results:
                if isinstance(result, list):
                    all_results.extend(result)
                elif isinstance(result, Exception):
                    logger.error(f"Search task exception: {str(result)}")
        
        # Remove duplicates
        seen_ids = set()
        unique_results = []
        for email in all_results:
            if email.get('id') and email['id'] not in seen_ids:
                seen_ids.add(email['id'])
                unique_results.append(email)
        
        # Sort by date
        unique_results.sort(key=lambda x: x.get('receivedDateTime', ''), reverse=True)
        
        return unique_results
    
    def build_simple_consensus(self, emails: List[Dict], query: str) -> str:
        """Build a simple, fast consensus without AI"""
        if not emails:
            return "No se encontraron correos relacionados con su búsqueda."
        
        # Group by date
        latest_email = emails[0] if emails else None
        
        consensus = f"**Resultados para: {query}**\n\n"
        consensus += f"Se encontraron {len(emails)} correos en los últimos {CURRENT_CONFIG['days_to_look_back']} días.\n\n"
        
        if latest_email:
            consensus += f"**Más reciente:**\n"
            consensus += f"- Fecha: {latest_email.get('receivedDateTime', '')[:10]}\n"
            consensus += f"- De: {latest_email.get('from', {}).get('emailAddress', {}).get('address', 'Desconocido')}\n"
            consensus += f"- Asunto: {latest_email.get('subject', 'Sin asunto')}\n\n"
        
        # Group emails by sender
        by_sender = {}
        for email in emails:
            sender = email.get('from', {}).get('emailAddress', {}).get('address', 'Desconocido')
            if sender not in by_sender:
                by_sender[sender] = []
            by_sender[sender].append(email)
        
        # Show top senders
        consensus += "**Principales remitentes:**\n"
        for sender, sender_emails in sorted(by_sender.items(), key=lambda x: len(x[1]), reverse=True)[:5]:
            consensus += f"- {sender}: {len(sender_emails)} correos\n"
        
        consensus += "\n**Correos recientes:**\n"
        for i, email in enumerate(emails[:10]):
            date = email.get('receivedDateTime', '')[:10]
            subject = email.get('subject', 'Sin asunto')
            consensus += f"{i+1}. [{date}] {subject}\n"
        
        return consensus
    
    async def build_ai_consensus(self, emails: List[Dict], query: str, search_terms: List[str]) -> str:
        """Build AI consensus"""
        if not emails:
            return "No se encontraron correos relacionados con su búsqueda."
        
        prompt = f"""Analiza estos {len(emails)} correos sobre: "{query}"

Términos buscados: {', '.join(search_terms)}
Período: últimos {CURRENT_CONFIG['days_to_look_back']} días

Proporciona una respuesta BREVE y DIRECTA con:
1. La información más importante encontrada (2-3 líneas)
2. Documentos o archivos clave mencionados
3. Última actualización o estado actual
4. Próximos pasos si están claros

Primeros 10 correos:
"""
        
        for i, email in enumerate(emails[:10]):
            subject = email.get('subject', 'Sin asunto')
            sender = email.get('from', {}).get('emailAddress', {}).get('address', 'Unknown')
            date = email.get('receivedDateTime', '')[:10]
            preview = email.get('bodyPreview', '')[:100]
            
            prompt += f"\n{i+1}. [{date}] {subject} - De: {sender}"
            if preview:
                prompt += f"\n   Vista previa: {preview}..."
        
        try:
            message = self.claude_client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=500,
                messages=[{"role": "user", "content": prompt}]
            )
            return message.content[0].text
        except Exception as e:
            logger.error(f"AI consensus error: {str(e)}")
            return self.build_simple_consensus(emails, query)
    
    async def search_emails_optimized(self, query: str, requested_by: str) -> Dict[str, Any]:
        """Main optimized search function"""
        start_time = time.time()
        
        # Check cache first
        if ENABLE_CACHE and query_cache:
            cache_key = query.lower().strip()
            if cache_key in query_cache:
                cached_time, cached_result = query_cache[cache_key]
                if (time.time() - cached_time) < CACHE_TTL_SECONDS:
                    logger.info(f"Returning cached result for: {query}")
                    cached_result['cached'] = True
                    cached_result['searchTime'] = "0.1 seconds (cached)"
                    return cached_result
        
        logger.info(f"Starting {SEARCH_MODE} search for '{query}' requested by {requested_by}")
        
        # Get smart search terms
        search_terms = self.get_smart_search_terms(query)
        logger.info(f"Using {len(search_terms)} search terms: {search_terms}")
        
        # Get active users with mailboxes
        users = await self.get_active_users()
        logger.info(f"Searching {len(users)} users: {users}")
        
        # Perform concurrent search
        emails = await self.concurrent_search(users, search_terms)
        logger.info(f"Found {len(emails)} total emails")
        
        # Build consensus
        if CURRENT_CONFIG['simple_consensus']:
            consensus = self.build_simple_consensus(emails, query)
        else:
            consensus = await self.build_ai_consensus(emails, query, search_terms)
        
        # Calculate search time
        search_time = round(time.time() - start_time, 1)
        
        # Prepare result
        result = {
            'success': True,
            'query': query,
            'searchMode': SEARCH_MODE,
            'searchTerms': search_terms,
            'stats': {
                'totalEmails': len(emails),
                'usersSearched': len(users),
                'daysSearched': CURRENT_CONFIG['days_to_look_back'],
                'searchTime': f"{search_time} seconds",
                'timestamp': datetime.now().isoformat()
            },
            'emails': emails[:50],  # Return top 50
            'consensus': consensus
        }
        
        # Cache result
        if ENABLE_CACHE and query_cache is not None:
            cache_key = query.lower().strip()
            query_cache[cache_key] = (time.time(), result)
        
        # Save to file
        filename = f"search_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{query[:20].replace(' ', '_')}.json"
        filepath = self.results_dir / filename
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=2, ensure_ascii=False)
        
        result['resultsFile'] = filename
        
        return result


# Initialize the service
search_service = OptimizedEmailSearchService()

# Routes
@app.route('/api/email-search', methods=['POST'])
def email_search():
    """Main email search endpoint"""
    try:
        data = request.json
        
        # Check authentication
        auth = data.get('auth', {})
        if not auth.get('authenticated') or auth.get('email') != 'admin@torelo.net':
            return jsonify({'error': 'Unauthorized'}), 401
        
        query = data.get('query', '').strip()
        if not query:
            return jsonify({'error': 'No search query provided'}), 400
        
        # Run optimized search
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        results = loop.run_until_complete(
            search_service.search_emails_optimized(query, auth.get('email'))
        )
        
        return jsonify(results)
        
    except Exception as e:
        logger.error(f"Search error: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy',
        'service': 'Torelo Email Search (Fixed)',
        'mode': SEARCH_MODE,
        'config': CURRENT_CONFIG,
        'timestamp': datetime.now().isoformat()
    })

@app.route('/api/search-config', methods=['GET'])
def get_search_config():
    """Get current search configuration"""
    return jsonify({
        'mode': SEARCH_MODE,
        'config': CURRENT_CONFIG,
        'priority_users': PRIORITY_USERS,
        'cache_enabled': ENABLE_CACHE,
        'spanish_priority': SPANISH_PRIORITY
    })

@app.route('/api/download-results/<filename>', methods=['GET'])
def download_results(filename):
    """Download full results file"""
    try:
        if not filename.endswith('.json'):
            return jsonify({'error': 'Invalid file type'}), 400
        
        filepath = Path('search_results') / filename
        if filepath.exists():
            return send_file(filepath, as_attachment=True)
        else:
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Error handlers
@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': 'Endpoint not found'}), 404

@app.errorhandler(500)
def server_error(e):
    return jsonify({'error': 'Internal server error'}), 500

if __name__ == '__main__':
    print(f"""
    ╔═══════════════════════════════════════╗
    ║   Torelo Email Search Server          ║
    ║   FIXED VERSION - {SEARCH_MODE} MODE        ║
    ║   Starting on http://0.0.0.0:5000     ║
    ╚═══════════════════════════════════════╝
    
    Configuration:
    - Mode: {SEARCH_MODE}
    - Search Terms: {CURRENT_CONFIG['max_search_terms']}
    - Users: {CURRENT_CONFIG['max_users_to_search']}
    - Days Back: {CURRENT_CONFIG['days_to_look_back']}
    - Concurrent: {CURRENT_CONFIG['concurrent_searches']}
    - Cache: {ENABLE_CACHE}
    
    Priority Users:
    {chr(10).join(f'  - {user}' for user in PRIORITY_USERS[:5])}
    
    Key Fixes:
    ✓ Corrected search syntax (no date in $search)
    ✓ Consolidated all config in Python
    ✓ Filter for users with mailboxes
    ✓ Better error handling
    """)
    
    # Run the Flask app - Debug OFF for production
    app.run(host='0.0.0.0', port=5000, debug=False)
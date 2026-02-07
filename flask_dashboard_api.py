"""
Flask API for Live Dashboard Updates
Real-time data refresh endpoint
"""

from flask import Flask, jsonify, send_file, request
from flask_cors import CORS
from template_generator import AdvancedTemplateGenerator
from api_modules.free_apis import FreeAPIHub
import os
from datetime import datetime
import threading

app = Flask(__name__)
CORS(app)

api_hub = FreeAPIHub()
current_file = 'ultimate_dashboard.xlsx'

@app.route('/')
def home():
    return jsonify({
        'name': 'Excel Template Creator API',
        'author': 'brendan-b25',
        'version': '1.0.0',
        'endpoints': {
            'POST /api/generate': 'Generate complete template',
            'GET /api/download': 'Download latest template',
            'POST /api/refresh/<sheet>': 'Refresh specific sheet data',
            'GET /api/data/<source>': 'Get live data from APIs',
            'GET /api/status': 'Get generation status'
        }
    })

@app.route('/api/generate', methods=['POST'])
def generate_template():
    """Generate a new complete template"""
    try:
        generator = AdvancedTemplateGenerator(current_file)
        filename = generator.generate_complete_template()
        
        return jsonify({
            'success': True,
            'filename': filename,
            'timestamp': datetime.now().isoformat(),
            'message': 'Template generated successfully'
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/download', methods=['GET'])
def download_template():
    """Download the generated template"""
    if os.path.exists(current_file):
        return send_file(
            current_file,
            as_attachment=True,
            download_name=f'dashboard_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
    return jsonify({'error': 'File not found'}), 404

@app.route('/api/data/weather', methods=['GET'])
def get_weather_data():
    """Get live weather data"""
    city = request.args.get('city', 'London')
    data = api_hub.get_weather(city)
    return jsonify(data)

@app.route('/api/data/crypto', methods=['GET'])
def get_crypto_data():
    """Get live crypto prices"""
    data = api_hub.get_crypto_prices()
    return jsonify(data)

@app.route('/api/data/exchange', methods=['GET'])
def get_exchange_data():
    """Get exchange rates"""
    base = request.args.get('base', 'USD')
    data = api_hub.get_exchange_rates(base)
    return jsonify(data)

@app.route('/api/data/github', methods=['GET'])
def get_github_data():
    """Get GitHub repo stats"""
    owner = request.args.get('owner', 'microsoft')
    repo = request.args.get('repo', 'vscode')
    data = api_hub.get_github_repo(owner, repo)
    return jsonify(data)

@app.route('/api/status', methods=['GET'])
def get_status():
    """Get API status"""
    return jsonify({
        'status': 'online',
        'timestamp': datetime.now().isoformat(),
        'file_exists': os.path.exists(current_file),
        'file_size': os.path.getsize(current_file) if os.path.exists(current_file) else 0,
        'apis_available': {
            'weather': True,
            'crypto': True,
            'exchange': True,
            'github': True,
            'users': True
        }
    })

if __name__ == '__main__':
    print("\n" + "="*60)
    print("🚀 Excel Template Creator API Server")
    print("="*60)
    print("\n📡 Server running at: http://localhost:5000")
    print("\n📚 Available endpoints:")
    print("   • POST http://localhost:5000/api/generate")
    print("   • GET  http://localhost:5000/api/download")
    print("   • GET  http://localhost:5000/api/data/weather?city=London")
    print("   • GET  http://localhost:5000/api/data/crypto")
    print("   • GET  http://localhost:5000/api/status")
    print("\n💡 Generate template: curl -X POST http://localhost:5000/api/generate")
    print("=" * 60 + "\n")
    
    app.run(debug=True, host='0.0.0.0', port=5000)

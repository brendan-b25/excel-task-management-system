#!/bin/bash
cd ~/excel-template-creator
source venv/bin/activate

# Start API in background
python3 flask_dashboard_api.py > api.log 2>&1 &
API_PID=$!

# Start web interface in background
python3 web_interface.py > web.log 2>&1 &
WEB_PID=$!

echo "🚀 Started!"
echo "   API (port 5000): PID $API_PID"
echo "   Web (port 8080): PID $WEB_PID"
echo ""
echo "🌐 Open: http://localhost:8080"
echo ""
echo "To stop: kill $API_PID $WEB_PID"

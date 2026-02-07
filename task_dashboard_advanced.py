"""
Advanced Live Web Dashboard with API Configuration
"""

from flask import Flask, render_template_string, jsonify, request, redirect
from flask_cors import CORS
from task_management_with_apis import TaskManagementSystem, APIIntegrations
from openpyxl import load_workbook
from datetime import datetime
import os

app = Flask(__name__)
CORS(app)

TASK_FILE = 'task_management_advanced.xlsx'

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Advanced Task Dashboard</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        .container { max-width: 1400px; margin: 0 auto; }
        .header {
            background: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }
        h1 { color: #2C3E50; font-size: 32px; margin-bottom: 10px; }
        .tabs {
            display: flex;
            gap: 10px;
            margin: 20px 0;
            border-bottom: 2px solid #ecf0f1;
        }
        .tab {
            padding: 15px 25px;
            cursor: pointer;
            border: none;
            background: none;
            font-size: 16px;
            color: #7f8c8d;
            border-bottom: 3px solid transparent;
            transition: all 0.3s;
        }
        .tab.active {
            color: #3498db;
            border-bottom-color: #3498db;
            font-weight: bold;
        }
        .tab:hover { color: #2980b9; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        .kpis {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 20px;
        }
        .kpi-card {
            background: white;
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            text-align: center;
        }
        .kpi-label { color: #7f8c8d; font-size: 14px; text-transform: uppercase; margin-bottom: 10px; }
        .kpi-value { color: #2C3E50; font-size: 36px; font-weight: bold; }
        .section {
            background: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }
        .section-title { color: #2C3E50; font-size: 24px; margin-bottom: 20px; }
        .task-table { width: 100%; border-collapse: collapse; }
        .task-table th {
            background: #3498db;
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
        }
        .task-table td { padding: 12px 15px; border-bottom: 1px solid #ecf0f1; }
        .task-table tr:hover { background: #f8f9fa; }
        .priority {
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: bold;
            display: inline-block;
        }
        .priority-critical { background: #c0392b; color: white; }
        .priority-high { background: #e74c3c; color: white; }
        .priority-medium { background: #f39c12; color: white; }
        .priority-low { background: #27ae60; color: white; }
        .status {
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: bold;
            display: inline-block;
        }
        .status-in-progress { background: #3498db; color: white; }
        .status-not-started { background: #95a5a6; color: white; }
        .status-completed { background: #27ae60; color: white; }
        .btn {
            padding: 12px 25px;
            border-radius: 8px;
            border: none;
            cursor: pointer;
            font-size: 16px;
            font-weight: bold;
            transition: all 0.3s;
        }
        .btn-primary { background: #3498db; color: white; }
        .btn-primary:hover { background: #2980b9; }
        .btn-success { background: #27ae60; color: white; }
        .btn-success:hover { background: #229954; }
        .form-group { margin-bottom: 20px; }
        .form-group label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
            color: #2C3E50;
        }
        .form-group input,
        .form-group select,
        .form-group textarea {
            width: 100%;
            padding: 10px;
            border: 2px solid #ecf0f1;
            border-radius: 5px;
            font-size: 14px;
        }
        .api-card {
            border: 2px solid #ecf0f1;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 15px;
        }
        .api-card h3 { color: #2C3E50; margin-bottom: 10px; }
        .api-status {
            display: inline-block;
            padding: 5px 15px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: bold;
            margin-left: 10px;
        }
        .api-enabled { background: #27ae60; color: white; }
        .api-disabled { background: #95a5a6; color: white; }
        .tip-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 25px;
            border-radius: 15px;
            margin-bottom: 20px;
        }
        .tip-card h3 { margin-bottom: 15px; }
        .weather-suggestion {
            background: #ecf0f1;
            padding: 20px;
            border-radius: 10px;
            margin-top: 15px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🚀 Advanced Task Management Dashboard</h1>
            <p style="color: #7f8c8d;">With 10+ API Integrations</p>
        </div>
        
        <div class="tabs">
            <button class="tab active" onclick="showTab('dashboard')">📊 Dashboard</button>
            <button class="tab" onclick="showTab('tasks')">📝 Tasks</button>
            <button class="tab" onclick="showTab('apis')">🔌 API Config</button>
            <button class="tab" onclick="showTab('insights')">💡 Insights</button>
        </div>
        
        <!-- Dashboard Tab -->
        <div id="dashboard-tab" class="tab-content active">
            <div class="kpis">
                <div class="kpi-card">
                    <div class="kpi-label">Total Tasks</div>
                    <div class="kpi-value" id="total-tasks">0</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Active Tasks</div>
                    <div class="kpi-value" id="active-tasks">0</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Completed</div>
                    <div class="kpi-value" id="completed-tasks">0</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Team Members</div>
                    <div class="kpi-value" id="team-members">0</div>
                </div>
            </div>
            
            <div class="tip-card">
                <h3>💡 Productivity Tip</h3>
                <p id="productivity-tip">Loading...</p>
                <p style="margin-top: 10px; font-size: 14px; opacity: 0.8;" id="tip-author"></p>
            </div>
            
            <div class="section">
                <h2 class="section-title">⚠️ Outstanding Tasks</h2>
                <table class="task-table" id="dashboard-tasks">
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>Task</th>
                            <th>Assigned To</th>
                            <th>Priority</th>
                            <th>Due Date</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
        </div>
        
        <!-- Tasks Tab -->
        <div id="tasks-tab" class="tab-content">
            <div class="section">
                <h2 class="section-title">Add New Task</h2>
                <button class="btn btn-success" onclick="showAddTaskForm()">+ Create Task</button>
                
                <div id="task-form" style="display: none; margin-top: 20px;">
                    <form onsubmit="createTask(event)">
                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
                            <div class="form-group">
                                <label>Task Name *</label>
                                <input type="text" name="name" required>
                            </div>
                            <div class="form-group">
                                <label>Category *</label>
                                <select name="category" required>
                                    <option>Development</option>
                                    <option>Design</option>
                                    <option>Marketing</option>
                                    <option>Sales</option>
                                    <option>Support</option>
                                    <option>Admin</option>
                                    <option>Planning</option>
                                    <option>Research</option>
                                </select>
                            </div>
                            <div class="form-group">
                                <label>Assigned To *</label>
                                <input type="text" name="assigned_to" required>
                            </div>
                            <div class="form-group">
                                <label>Priority</label>
                                <select name="priority">
                                    <option>Low</option>
                                    <option selected>Medium</option>
                                    <option>High</option>
                                    <option>Critical</option>
                                </select>
                            </div>
                            <div class="form-group">
                                <label>Due Date</label>
                                <input type="date" name="due_date" required>
                            </div>
                            <div class="form-group">
                                <label>Status</label>
                                <select name="status">
                                    <option selected>Not Started</option>
                                    <option>In Progress</option>
                                    <option>On Hold</option>
                                </select>
                            </div>
                        </div>
                        <div class="form-group">
                            <label>Description</label>
                            <textarea name="description" rows="3"></textarea>
                        </div>
                        <div class="form-group">
                            <label>
                                <input type="checkbox" name="send_notifications" checked>
                                Send API Notifications (Slack, Discord, Telegram, etc.)
                            </label>
                        </div>
                        <button type="submit" class="btn btn-success">Create Task</button>
                        <button type="button" class="btn btn-primary" onclick="hideTaskForm()">Cancel</button>
                    </form>
                </div>
            </div>
            
            <div class="section">
                <h2 class="section-title">All Tasks</h2>
                <table class="task-table" id="all-tasks">
                    <thead>
                        <tr>
                            <th>ID</th>
                            <th>Task</th>
                            <th>Category</th>
                            <th>Assigned To</th>
                            <th>Priority</th>
                            <th>Due Date</th>
                            <th>Status</th>
                            <th>Progress</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
        </div>
        
        <!-- API Config Tab -->
        <div id="apis-tab" class="tab-content">
            <div class="section">
                <h2 class="section-title">🔌 Configure API Integrations</h2>
                
                <div class="api-card">
                    <h3>Slack <span class="api-status api-disabled" id="slack-status">Not Configured</span></h3>
                    <p>Send task notifications to Slack channels via webhooks.</p>
                    <div class="form-group">
                        <label>Webhook URL</label>
                        <input type="text" id="slack-webhook" placeholder="https://hooks.slack.com/services/...">
                        <button class="btn btn-primary" onclick="saveAPI('slack')">Save</button>
                        <button class="btn btn-success" onclick="testAPI('slack')">Test</button>
                    </div>
                    <p style="margin-top: 10px; font-size: 14px; color: #7f8c8d;">
                        Get webhook: <a href="https://api.slack.com/messaging/webhooks" target="_blank">https://api.slack.com/messaging/webhooks</a>
                    </p>
                </div>
                
                <div class="api-card">
                    <h3>Discord <span class="api-status api-disabled" id="discord-status">Not Configured</span></h3>
                    <p>Send task notifications to Discord channels.</p>
                    <div class="form-group">
                        <label>Webhook URL</label>
                        <input type="text" id="discord-webhook" placeholder="https://discord.com/api/webhooks/...">
                        <button class="btn btn-primary" onclick="saveAPI('discord')">Save</button>
                        <button class="btn btn-success" onclick="testAPI('discord')">Test</button>
                    </div>
                    <p style="margin-top: 10px; font-size: 14px; color: #7f8c8d;">
                        Create webhook in Discord Server Settings → Integrations
                    </p>
                </div>
                
                <div class="api-card">
                    <h3>Telegram <span class="api-status api-disabled" id="telegram-status">Not Configured</span></h3>
                    <p>Send task notifications via Telegram bot.</p>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px;">
                        <div class="form-group">
                            <label>Bot Token</label>
                            <input type="text" id="telegram-token" placeholder="123456:ABC-DEF...">
                        </div>
                        <div class="form-group">
                            <label>Chat ID</label>
                            <input type="text" id="telegram-chat" placeholder="-1001234567890">
                        </div>
                    </div>
                    <button class="btn btn-primary" onclick="saveAPI('telegram')">Save</button>
                    <button class="btn btn-success" onclick="testAPI('telegram')">Test</button>
                    <p style="margin-top: 10px; font-size: 14px; color: #7f8c8d;">
                        Create bot: <a href="https://t.me/BotFather" target="_blank">@BotFather</a>
                    </p>
                </div>
                
                <div class="api-card">
                    <h3>GitHub Issues <span class="api-status api-disabled" id="github-status">Not Configured</span></h3>
                    <p>Automatically create GitHub issues from tasks.</p>
                    <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px;">
                        <div class="form-group">
                            <label>Token</label>
                            <input type="password" id="github-token" placeholder="ghp_...">
                        </div>
                        <div class="form-group">
                            <label>Owner</label>
                            <input type="text" id="github-owner" placeholder="username">
                        </div>
                        <div class="form-group">
                            <label>Repository</label>
                            <input type="text" id="github-repo" placeholder="repo-name">
                        </div>
                    </div>
                    <button class="btn btn-primary" onclick="saveAPI('github')">Save</button>
                    <button class="btn btn-success" onclick="testAPI('github')">Test</button>
                    <p style="margin-top: 10px; font-size: 14px; color: #7f8c8d;">
                        Create token: <a href="https://github.com/settings/tokens" target="_blank">GitHub Settings → Tokens</a>
                    </p>
                </div>
            </div>
        </div>
        
        <!-- Insights Tab -->
        <div id="insights-tab" class="tab-content">
            <div class="section">
                <h2 class="section-title">🌤️ Weather-Based Deadline Suggestions</h2>
                <div class="form-group">
                    <label>City</label>
                    <input type="text" id="weather-city" value="London" placeholder="Enter city name">
                    <button class="btn btn-primary" onclick="getWeatherSuggestion()">Get Suggestion</button>
                </div>
                <div id="weather-result" class="weather-suggestion" style="display: none;"></div>
            </div>
            
            <div class="section">
                <h2 class="section-title">📊 Task Analytics</h2>
                <div id="analytics-content">Loading analytics...</div>
            </div>
            
            <div class="section">
                <h2 class="section-title">📥 Download Excel File</h2>
                <p>Download the complete Excel file with all tasks and integrations.</p>
                <a href="/api/download" class="btn btn-success" style="display: inline-block; text-decoration: none; margin-top: 10px;">📥 Download Excel</a>
            </div>
        </div>
    </div>
    
    <script>
        let apiConfig = {};
        
        function showTab(tabName) {
            document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            event.target.classList.add('active');
            document.getElementById(tabName + '-tab').classList.add('active');
            
            if (tabName === 'dashboard') loadDashboard();
            if (tabName === 'tasks') loadAllTasks();
            if (tabName === 'apis') loadAPIConfig();
            if (tabName === 'insights') loadAnalytics();
        }
        
        async function loadDashboard() {
            try {
                const response = await fetch('/api/dashboard');
                const data = await response.json();
                
                document.getElementById('total-tasks').textContent = data.kpis.total;
                document.getElementById('active-tasks').textContent = data.kpis.active;
                document.getElementById('completed-tasks').textContent = data.kpis.completed;
                document.getElementById('team-members').textContent = data.kpis.team_members;
                
                document.getElementById('productivity-tip').textContent = data.productivity_tip.tip;
                document.getElementById('tip-author').textContent = '— ' + data.productivity_tip.author;
                
                const tbody = document.querySelector('#dashboard-tasks tbody');
                tbody.innerHTML = '';
                data.outstanding_tasks.forEach(task => {
                    tbody.innerHTML += `
                        <tr>
                            <td>${task.id}</td>
                            <td><strong>${task.name}</strong></td>
                            <td>${task.assigned_to}</td>
                            <td><span class="priority priority-${task.priority.toLowerCase()}">${task.priority}</span></td>
                            <td>${task.due_date}</td>
                            <td><span class="status status-${task.status.toLowerCase().replace(' ', '-')}">${task.status}</span></td>
                        </tr>
                    `;
                });
            } catch (error) {
                console.error('Error loading dashboard:', error);
            }
        }
        
        async function loadAllTasks() {
            try {
                const response = await fetch('/api/tasks');
                const data = await response.json();
                
                const tbody = document.querySelector('#all-tasks tbody');
                tbody.innerHTML = '';
                data.tasks.forEach(task => {
                    tbody.innerHTML += `
                        <tr>
                            <td>${task.id}</td>
                            <td><strong>${task.name}</strong></td>
                            <td>${task.category}</td>
                            <td>${task.assigned_to}</td>
                            <td><span class="priority priority-${task.priority.toLowerCase()}">${task.priority}</span></td>
                            <td>${task.due_date}</td>
                            <td><span class="status status-${task.status.toLowerCase().replace(' ', '-')}">${task.status}</span></td>
                            <td>${task.progress}%</td>
                        </tr>
                    `;
                });
            } catch (error) {
                console.error('Error loading tasks:', error);
            }
        }
        
        function showAddTaskForm() {
            document.getElementById('task-form').style.display = 'block';
        }
        
        function hideTaskForm() {
            document.getElementById('task-form').style.display = 'none';
        }
        
        async function createTask(event) {
            event.preventDefault();
            const formData = new FormData(event.target);
            const taskData = Object.fromEntries(formData);
            
            try {
                const response = await fetch('/api/tasks', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify(taskData)
                });
                
                const result = await response.json();
                if (result.success) {
                    alert('✅ Task created successfully!');
                    hideTaskForm();
                    event.target.reset();
                    loadAllTasks();
                    loadDashboard();
                } else {
                    alert('❌ Error: ' + result.error);
                }
            } catch (error) {
                alert('Error: ' + error.message);
            }
        }
        
        function loadAPIConfig() {
            // Load saved API configs from localStorage
            apiConfig = JSON.parse(localStorage.getItem('apiConfig') || '{}');
            
            if (apiConfig.slack_webhook) {
                document.getElementById('slack-webhook').value = apiConfig.slack_webhook;
                document.getElementById('slack-status').textContent = 'Configured';
                document.getElementById('slack-status').className = 'api-status api-enabled';
            }
            
            if (apiConfig.discord_webhook) {
                document.getElementById('discord-webhook').value = apiConfig.discord_webhook;
                document.getElementById('discord-status').textContent = 'Configured';
                document.getElementById('discord-status').className = 'api-status api-enabled';
            }
            
            // Add more configs...
        }
        
        function saveAPI(service) {
            if (service === 'slack') {
                apiConfig.slack_webhook = document.getElementById('slack-webhook').value;
            } else if (service === 'discord') {
                apiConfig.discord_webhook = document.getElementById('discord-webhook').value;
            } else if (service === 'telegram') {
                apiConfig.telegram_token = document.getElementById('telegram-token').value;
                apiConfig.telegram_chat = document.getElementById('telegram-chat').value;
            } else if (service === 'github') {
                apiConfig.github_token = document.getElementById('github-token').value;
                apiConfig.github_owner = document.getElementById('github-owner').value;
                apiConfig.github_repo = document.getElementById('github-repo').value;
            }
            
            localStorage.setItem('apiConfig', JSON.stringify(apiConfig));
            alert('✅ API configuration saved!');
            loadAPIConfig();
        }
        
        async function testAPI(service) {
            const testTask = {
                name: 'Test Task',
                assigned_to: 'Test User',
                priority: 'Medium',
                category: 'Testing',
                due_date: new Date().toISOString().split('T')[0],
                description: 'This is a test notification'
            };
            
            try {
                const response = await fetch('/api/test-notification', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({service, task: testTask, config: apiConfig})
                });
                
                const result = await response.json();
                if (result.success) {
                    alert('✅ Test notification sent successfully!');
                } else {
                    alert('❌ Test failed: ' + result.error);
                }
            } catch (error) {
                alert('Error: ' + error.message);
            }
        }
        
        async function getWeatherSuggestion() {
            const city = document.getElementById('weather-city').value;
            try {
                const response = await fetch(`/api/weather-suggestion?city=${city}`);
                const data = await response.json();
                
                const resultDiv = document.getElementById('weather-result');
                resultDiv.innerHTML = `
                    <h3>✅ Suggested Deadline: ${data.suggested_date}</h3>
                    <p><strong>Reason:</strong> ${data.reason}</p>
                    <h4 style="margin-top: 15px;">Alternative Dates:</h4>
                    <ul>
                        ${data.alternatives.map(alt => `
                            <li>${alt.date} - ${alt.weather}, ${alt.temp}°C (Score: ${alt.score}/10)</li>
                        `).join('')}
                    </ul>
                `;
                resultDiv.style.display = 'block';
            } catch (error) {
                alert('Error getting weather suggestion: ' + error.message);
            }
        }
        
        async function loadAnalytics() {
            try {
                const response = await fetch('/api/analytics');
                const data = await response.json();
                
                const content = document.getElementById('analytics-content');
                content.innerHTML = `
                    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px;">
                        <div>
                            <h3>By Category</h3>
                            <ul>
                                ${Object.entries(data.by_category).map(([cat, count]) => `
                                    <li>${cat}: ${count} tasks</li>
                                `).join('')}
                            </ul>
                        </div>
                        <div>
                            <h3>By Priority</h3>
                            <ul>
                                ${Object.entries(data.by_priority).map(([pri, count]) => `
                                    <li>${pri}: ${count} tasks</li>
                                `).join('')}
                            </ul>
                        </div>
                        <div>
                            <h3>By Status</h3>
                            <ul>
                                ${Object.entries(data.by_status).map(([status, count]) => `
                                    <li>${status}: ${count} tasks</li>
                                `).join('')}
                            </ul>
                        </div>
                    </div>
                `;
            } catch (error) {
                console.error('Error loading analytics:', error);
            }
        }
        
        // Load dashboard on page load
        window.onload = () => {
            loadDashboard();
        };
        
        // Auto-refresh every 60 seconds
        setInterval(loadDashboard, 60000);
    </script>
</body>
</html>
'''

@app.route('/')
def home():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/dashboard')
def get_dashboard():
    try:
        if not os.path.exists(TASK_FILE):
            return jsonify({'error': 'File not found'}), 404
        
        wb = load_workbook(TASK_FILE, data_only=True)
        ws = wb['Task Entry']
        
        tasks = []
        team_members = set()
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:
                tasks.append({
                    'id': row[0],
                    'name': row[1],
                    'category': row[2],
                    'assigned_to': row[3],
                    'priority': row[4],
                    'due_date': str(row[5]) if row[5] else '',
                    'status': row[6],
                    'progress': row[7] or 0
                })
                if row[3]:
                    team_members.add(row[3])
        
        total = len(tasks)
        active = sum(1 for t in tasks if t['status'] in ['In Progress', 'Not Started'])
        completed = sum(1 for t in tasks if t['status'] == 'Completed')
        outstanding = [t for t in tasks if t['status'] not in ['Completed', 'Cancelled']]
        
        # Get productivity tip
        api = APIIntegrations()
        tip = api.get_productivity_tip()
        
        return jsonify({
            'kpis': {
                'total': total,
                'active': active,
                'completed': completed,
                'team_members': len(team_members)
            },
            'outstanding_tasks': outstanding[:10],
            'productivity_tip': tip
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/tasks', methods=['GET'])
def get_tasks():
    try:
        wb = load_workbook(TASK_FILE, data_only=True)
        ws = wb['Task Entry']
        
        tasks = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:
                tasks.append({
                    'id': row[0],
                    'name': row[1],
                    'category': row[2],
                    'assigned_to': row[3],
                    'priority': row[4],
                    'due_date': str(row[5]) if row[5] else '',
                    'status': row[6],
                    'progress': row[7] or 0,
                    'description': row[8] or ''
                })
        
        return jsonify({'tasks': tasks})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/tasks', methods=['POST'])
def create_task():
    try:
        task_data = request.json
        notify = task_data.get('send_notifications', True)
        
        tms = TaskManagementSystem(TASK_FILE)
        task_id = tms.add_task(task_data, notify=notify)
        
        return jsonify({'success': True, 'task_id': task_id})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/weather-suggestion')
def weather_suggestion():
    try:
        city = request.args.get('city', 'London')
        api = APIIntegrations()
        suggestion = api.suggest_deadline_based_on_weather(city)
        return jsonify(suggestion)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/analytics')
def get_analytics():
    try:
        wb = load_workbook(TASK_FILE, data_only=True)
        ws = wb['Task Entry']
        
        by_category = {}
        by_priority = {}
        by_status = {}
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:
                by_category[row[2]] = by_category.get(row[2], 0) + 1
                by_priority[row[4]] = by_priority.get(row[4], 0) + 1
                by_status[row[6]] = by_status.get(row[6], 0) + 1
        
        return jsonify({
            'by_category': by_category,
            'by_priority': by_priority,
            'by_status': by_status
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/test-notification', methods=['POST'])
def test_notification():
    try:
        data = request.json
        service = data['service']
        task = data['task']
        config = data['config']
        
        api = APIIntegrations()
        
        success = False
        if service == 'slack' and config.get('slack_webhook'):
            success = api.send_slack_notification(config['slack_webhook'], task)
        elif service == 'discord' and config.get('discord_webhook'):
            success = api.send_discord_notification(config['discord_webhook'], task)
        elif service == 'telegram' and config.get('telegram_token'):
            success = api.send_telegram_message(
                config['telegram_token'],
                config['telegram_chat'],
                task
            )
        
        return jsonify({'success': success})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/download')
def download():
    if os.path.exists(TASK_FILE):
        from flask import send_file
        return send_file(TASK_FILE, as_attachment=True)
    return "File not found", 404

if __name__ == '__main__':
    print("\n" + "="*80)
    print("🚀 ADVANCED TASK MANAGEMENT DASHBOARD WITH API INTEGRATIONS")
    print("="*80)
    print("\n📱 Open in your browser:")
    print("   → http://localhost:9000")
    print("\n🎯 Features:")
    print("   • Live dashboard with real-time KPIs")
    print("   • Add/manage tasks from web interface")
    print("   • Configure 10+ API integrations")
    print("   • Weather-based deadline suggestions")
    print("   • Productivity tips")
    print("   • Task analytics")
    print("   • Auto-refresh every 60 seconds")
    print("\n🔌 Supported APIs:")
    print("   • Slack, Discord, Telegram")
    print("   • GitHub Issues")
    print("   • Trello, Notion, Airtable")
    print("   • Weather API")
    print("\n" + "="*80 + "\n")
    
    app.run(debug=True, host='0.0.0.0', port=9000)

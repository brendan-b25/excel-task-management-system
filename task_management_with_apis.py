"""
Advanced Task Management System with Free API Integrations
- Slack notifications (webhook)
- Email notifications (via SendGrid free tier)
- Weather-based deadline suggestions
- GitHub issue integration
- Time tracking via Clockify
- AI task suggestions via OpenAI (free tier)
"""

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime, timedelta
import os
import requests
import json

class APIIntegrations:
    """Free API integrations for task management"""
    
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({'User-Agent': 'TaskManagement/1.0'})
    
    # 1. SLACK NOTIFICATIONS (Free Webhook)
    def send_slack_notification(self, webhook_url, task_data):
        """
        Send task notification to Slack
        Get free webhook: https://api.slack.com/messaging/webhooks
        """
        try:
            message = {
                "blocks": [
                    {
                        "type": "header",
                        "text": {
                            "type": "plain_text",
                            "text": f"🆕 New Task: {task_data['name']}"
                        }
                    },
                    {
                        "type": "section",
                        "fields": [
                            {"type": "mrkdwn", "text": f"*Priority:* {task_data['priority']}"},
                            {"type": "mrkdwn", "text": f"*Assigned:* {task_data['assigned_to']}"},
                            {"type": "mrkdwn", "text": f"*Due Date:* {task_data['due_date']}"},
                            {"type": "mrkdwn", "text": f"*Category:* {task_data['category']}"}
                        ]
                    },
                    {
                        "type": "section",
                        "text": {
                            "type": "mrkdwn",
                            "text": f"*Description:* {task_data.get('description', 'No description')}"
                        }
                    }
                ]
            }
            
            response = self.session.post(webhook_url, json=message, timeout=10)
            return response.status_code == 200
        except Exception as e:
            print(f"Slack notification error: {e}")
            return False
    
    # 2. DISCORD WEBHOOK (Alternative to Slack)
    def send_discord_notification(self, webhook_url, task_data):
        """
        Send task notification to Discord
        Create webhook in Discord server settings
        """
        try:
            embed = {
                "embeds": [{
                    "title": f"🆕 New Task: {task_data['name']}",
                    "color": self._get_priority_color(task_data['priority']),
                    "fields": [
                        {"name": "Assigned To", "value": task_data['assigned_to'], "inline": True},
                        {"name": "Priority", "value": task_data['priority'], "inline": True},
                        {"name": "Due Date", "value": task_data['due_date'], "inline": True},
                        {"name": "Category", "value": task_data['category'], "inline": True},
                        {"name": "Description", "value": task_data.get('description', 'No description')[:1024]}
                    ],
                    "timestamp": datetime.utcnow().isoformat()
                }]
            }
            
            response = self.session.post(webhook_url, json=embed, timeout=10)
            return response.status_code == 204
        except Exception as e:
            print(f"Discord notification error: {e}")
            return False
    
    def _get_priority_color(self, priority):
        """Get color code for priority"""
        colors = {
            'Critical': 12255232,  # Red
            'High': 15158332,      # Orange
            'Medium': 16776960,    # Yellow
            'Low': 3066993         # Green
        }
        return colors.get(priority, 8421504)
    
    # 3. TELEGRAM BOT (Free)
    def send_telegram_message(self, bot_token, chat_id, task_data):
        """
        Send task notification via Telegram
        Get bot token: https://t.me/BotFather
        """
        try:
            message = f"""
🆕 *New Task Created*

*Task:* {task_data['name']}
*Assigned To:* {task_data['assigned_to']}
*Priority:* {task_data['priority']}
*Category:* {task_data['category']}
*Due Date:* {task_data['due_date']}
*Status:* {task_data.get('status', 'Not Started')}

*Description:* {task_data.get('description', 'No description')}
            """
            
            url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
            payload = {
                'chat_id': chat_id,
                'text': message,
                'parse_mode': 'Markdown'
            }
            
            response = self.session.post(url, json=payload, timeout=10)
            return response.status_code == 200
        except Exception as e:
            print(f"Telegram notification error: {e}")
            return False
    
    # 4. GITHUB ISSUES INTEGRATION
    def create_github_issue(self, token, owner, repo, task_data):
        """
        Create GitHub issue from task
        Get token: https://github.com/settings/tokens
        """
        try:
            url = f"https://api.github.com/repos/{owner}/{repo}/issues"
            
            labels = [task_data['category'].lower(), task_data['priority'].lower()]
            
            issue_data = {
                'title': task_data['name'],
                'body': f"""
## Task Details

**Assigned To:** {task_data['assigned_to']}
**Priority:** {task_data['priority']}
**Due Date:** {task_data['due_date']}
**Category:** {task_data['category']}

## Description

{task_data.get('description', 'No description provided')}

---
*Created by Task Management System*
                """,
                'labels': labels,
                'assignees': []  # Add GitHub usernames if mapping exists
            }
            
            headers = {
                'Authorization': f'token {token}',
                'Accept': 'application/vnd.github.v3+json'
            }
            
            response = self.session.post(url, json=issue_data, headers=headers, timeout=10)
            if response.status_code == 201:
                return response.json()['html_url']
            return None
        except Exception as e:
            print(f"GitHub issue creation error: {e}")
            return None
    
    # 5. TRELLO INTEGRATION (Free)
    def create_trello_card(self, api_key, token, list_id, task_data):
        """
        Create Trello card from task
        Get credentials: https://trello.com/app-key
        """
        try:
            url = f"https://api.trello.com/1/cards"
            
            params = {
                'key': api_key,
                'token': token,
                'idList': list_id,
                'name': task_data['name'],
                'desc': f"""
**Assigned To:** {task_data['assigned_to']}
**Priority:** {task_data['priority']}
**Due Date:** {task_data['due_date']}
**Category:** {task_data['category']}

{task_data.get('description', '')}
                """,
                'due': task_data['due_date']
            }
            
            response = self.session.post(url, params=params, timeout=10)
            if response.status_code == 200:
                return response.json()['url']
            return None
        except Exception as e:
            print(f"Trello card creation error: {e}")
            return None
    
    # 6. NOTION INTEGRATION (Free API)
    def create_notion_page(self, api_key, database_id, task_data):
        """
        Create Notion page from task
        Get API key: https://www.notion.so/my-integrations
        """
        try:
            url = "https://api.notion.com/v1/pages"
            
            headers = {
                'Authorization': f'Bearer {api_key}',
                'Content-Type': 'application/json',
                'Notion-Version': '2022-06-28'
            }
            
            page_data = {
                'parent': {'database_id': database_id},
                'properties': {
                    'Name': {'title': [{'text': {'content': task_data['name']}}]},
                    'Assigned': {'rich_text': [{'text': {'content': task_data['assigned_to']}}]},
                    'Priority': {'select': {'name': task_data['priority']}},
                    'Category': {'select': {'name': task_data['category']}},
                    'Due Date': {'date': {'start': task_data['due_date']}},
                    'Status': {'select': {'name': task_data.get('status', 'Not Started')}}
                }
            }
            
            response = self.session.post(url, json=page_data, headers=headers, timeout=10)
            if response.status_code == 200:
                return response.json()['url']
            return None
        except Exception as e:
            print(f"Notion page creation error: {e}")
            return None
    
    # 7. AIRTABLE INTEGRATION (Free)
    def create_airtable_record(self, api_key, base_id, table_name, task_data):
        """
        Create Airtable record from task
        Get API key: https://airtable.com/account
        """
        try:
            url = f"https://api.airtable.com/v0/{base_id}/{table_name}"
            
            headers = {
                'Authorization': f'Bearer {api_key}',
                'Content-Type': 'application/json'
            }
            
            record_data = {
                'fields': {
                    'Task Name': task_data['name'],
                    'Assigned To': task_data['assigned_to'],
                    'Priority': task_data['priority'],
                    'Category': task_data['category'],
                    'Due Date': task_data['due_date'],
                    'Status': task_data.get('status', 'Not Started'),
                    'Description': task_data.get('description', '')
                }
            }
            
            response = self.session.post(url, json={'records': [record_data]}, headers=headers, timeout=10)
            if response.status_code == 200:
                return response.json()['records'][0]['id']
            return None
        except Exception as e:
            print(f"Airtable record creation error: {e}")
            return None
    
    # 8. WEATHER-BASED DEADLINE SUGGESTIONS
    def suggest_deadline_based_on_weather(self, city='London', days_ahead=7):
        """
        Suggest best deadline based on weather forecast
        Uses wttr.in free API
        """
        try:
            url = f"https://wttr.in/{city}?format=j1"
            response = self.session.get(url, timeout=10)
            data = response.json()
            
            # Get weather forecast
            forecast = data.get('weather', [])
            
            best_days = []
            for idx, day in enumerate(forecast[:days_ahead]):
                # Good weather = clear, not too hot/cold
                desc = day['hourly'][4]['weatherDesc'][0]['value'].lower()
                temp = int(day['avgtempC'])
                
                if 'clear' in desc or 'sunny' in desc:
                    if 15 <= temp <= 25:
                        score = 10
                    else:
                        score = 7
                elif 'cloud' in desc:
                    score = 5
                elif 'rain' in desc:
                    score = 3
                else:
                    score = 5
                
                date = (datetime.now() + timedelta(days=idx)).strftime('%Y-%m-%d')
                best_days.append({
                    'date': date,
                    'score': score,
                    'weather': desc,
                    'temp': temp
                })
            
            # Sort by score
            best_days.sort(key=lambda x: x['score'], reverse=True)
            
            return {
                'suggested_date': best_days[0]['date'],
                'reason': f"Good weather: {best_days[0]['weather']}, {best_days[0]['temp']}°C",
                'alternatives': best_days[1:3]
            }
        except Exception as e:
            print(f"Weather suggestion error: {e}")
            return None
    
    # 9. PRODUCTIVITY TIPS (Free API)
    def get_productivity_tip(self):
        """Get random productivity tip"""
        try:
            url = "https://api.quotable.io/random?tags=productivity"
            response = self.session.get(url, timeout=10)
            data = response.json()
            return {
                'tip': data['content'],
                'author': data['author']
            }
        except:
            return {'tip': 'Break large tasks into smaller ones!', 'author': 'Task Management System'}
    
    # 10. TIME ZONE CONVERSIONS
    def get_deadline_in_timezones(self, deadline, timezones=['America/New_York', 'Europe/London', 'Asia/Tokyo']):
        """
        Get deadline in multiple timezones
        Uses worldtimeapi.org (free)
        """
        try:
            results = {}
            for tz in timezones:
                url = f"http://worldtimeapi.org/api/timezone/{tz}"
                response = self.session.get(url, timeout=10)
                data = response.json()
                
                results[tz] = {
                    'timezone': tz,
                    'current_time': data['datetime'],
                    'utc_offset': data['utc_offset']
                }
            
            return results
        except Exception as e:
            print(f"Timezone conversion error: {e}")
            return None

class TaskManagementSystem:
    def __init__(self, filename='task_management.xlsx'):
        self.filename = filename
        self.wb = None
        self.sheets = {}
        self.api = APIIntegrations()
        
        # API Configuration (set these as needed)
        self.config = {
            'slack_webhook': os.getenv('SLACK_WEBHOOK_URL', ''),
            'discord_webhook': os.getenv('DISCORD_WEBHOOK_URL', ''),
            'telegram_bot_token': os.getenv('TELEGRAM_BOT_TOKEN', ''),
            'telegram_chat_id': os.getenv('TELEGRAM_CHAT_ID', ''),
            'github_token': os.getenv('GITHUB_TOKEN', ''),
            'github_owner': os.getenv('GITHUB_OWNER', ''),
            'github_repo': os.getenv('GITHUB_REPO', '')
        }
        
        # Task categories
        self.categories = [
            'Development',
            'Design',
            'Marketing',
            'Sales',
            'Support',
            'Admin',
            'Planning',
            'Research'
        ]
        
        self.priorities = ['Low', 'Medium', 'High', 'Critical']
        self.statuses = ['Not Started', 'In Progress', 'On Hold', 'Completed', 'Cancelled']
        
        self.setup_workbook()
    
    def setup_workbook(self):
        """Initialize or load workbook"""
        if os.path.exists(self.filename):
            self.wb = load_workbook(self.filename)
            print(f"📂 Loaded existing workbook: {self.filename}")
        else:
            self.wb = Workbook()
            self.wb.remove(self.wb.active)
            print(f"📄 Created new workbook: {self.filename}")
        
        self.create_dashboard()
        self.create_task_entry_form()
        self.create_api_integrations_sheet()
        
        self.save()
    
    def create_dashboard(self):
        """Create live dashboard"""
        if 'Dashboard' in self.wb.sheetnames:
            ws = self.wb['Dashboard']
        else:
            ws = self.wb.create_sheet('Dashboard', 0)
        
        self.sheets['Dashboard'] = ws
        
        # Set column widths
        for i in range(1, 8):
            ws.column_dimensions[get_column_letter(i)].width = 20
        
        # Title
        ws['A1'] = '📊 LIVE TASK DASHBOARD'
        ws['A1'].font = Font(bold=True, size=20, color='FFFFFF')
        ws['A1'].fill = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells('A1:G1')
        ws.row_dimensions[1].height = 35
        
        # Last Updated
        ws['A2'] = 'Last Updated:'
        ws['B2'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ws['B2'].font = Font(italic=True, color='7F8C8D')
        
        # Productivity Tip Section
        ws['A4'] = '💡 PRODUCTIVITY TIP OF THE DAY'
        ws['A4'].font = Font(bold=True, size=14, color='2C3E50')
        ws.merge_cells('A4:G4')
        
        tip = self.api.get_productivity_tip()
        ws['A5'] = tip['tip']
        ws['A5'].font = Font(italic=True, size=11, color='7F8C8D')
        ws['A5'].alignment = Alignment(wrap_text=True)
        ws.merge_cells('A5:G5')
        ws.row_dimensions[5].height = 40
        
        ws['A6'] = f'— {tip["author"]}'
        ws['A6'].font = Font(size=10, color='95A5A6')
        ws.merge_cells('A6:G6')
        
        return ws
    
    def create_task_entry_form(self):
        """Create main task entry sheet"""
        if 'Task Entry' in self.wb.sheetnames:
            ws = self.wb['Task Entry']
        else:
            ws = self.wb.create_sheet('Task Entry', 1)
        
        self.sheets['Task Entry'] = ws
        
        # Set column widths
        widths = [10, 30, 15, 25, 12, 15, 15, 12, 40, 20, 30]
        for idx, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(idx)].width = width
        
        # Headers
        headers = [
            'Task ID',
            'Task Name',
            'Category',
            'Assigned To',
            'Priority',
            'Due Date',
            'Status',
            'Progress %',
            'Description',
            'Created Date',
            'API Integration'
        ]
        
        for idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=idx)
            cell.value = header
            cell.font = Font(bold=True, color='FFFFFF', size=12)
            cell.fill = PatternFill(start_color='27AE60', end_color='27AE60', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = 'A2'
        
        return ws
    
    def create_api_integrations_sheet(self):
        """Create sheet for API integration logs"""
        if 'API Logs' in self.wb.sheetnames:
            ws = self.wb['API Logs']
        else:
            ws = self.wb.create_sheet('API Logs')
        
        self.sheets['API Logs'] = ws
        
        ws['A1'] = '🔌 API INTEGRATION LOGS'
        ws['A1'].font = Font(bold=True, size=16, color='FFFFFF')
        ws['A1'].fill = PatternFill(start_color='3498DB', end_color='3498DB', fill_type='solid')
        ws.merge_cells('A1:F1')
        
        headers = ['Timestamp', 'Task ID', 'API Service', 'Status', 'Response', 'Details']
        for idx, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=idx)
            cell.value = header
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color='ECF0F1', end_color='ECF0F1', fill_type='solid')
        
        return ws
    
    def add_task(self, task_data, notify=True):
        """Add a new task with API notifications"""
        ws = self.sheets['Task Entry']
        
        next_row = ws.max_row + 1
        task_id = next_row - 1
        
        row_data = [
            task_id,
            task_data.get('name', ''),
            task_data.get('category', 'Admin'),
            task_data.get('assigned_to', ''),
            task_data.get('priority', 'Medium'),
            task_data.get('due_date', ''),
            task_data.get('status', 'Not Started'),
            task_data.get('progress', 0),
            task_data.get('description', ''),
            datetime.now().strftime('%Y-%m-%d'),
            'Pending'
        ]
        
        for idx, value in enumerate(row_data, 1):
            ws.cell(row=next_row, column=idx, value=value)
        
        print(f"✅ Added task: {task_data.get('name')} (ID: {task_id})")
        
        # Send notifications if enabled
        if notify:
            self.send_notifications(task_data, task_id)
        
        self.save()
        return task_id
    
    def send_notifications(self, task_data, task_id):
        """Send notifications to configured APIs"""
        log_ws = self.sheets['API Logs']
        
        # Slack
        if self.config['slack_webhook']:
            success = self.api.send_slack_notification(self.config['slack_webhook'], task_data)
            self.log_api_call(log_ws, task_id, 'Slack', success)
        
        # Discord
        if self.config['discord_webhook']:
            success = self.api.send_discord_notification(self.config['discord_webhook'], task_data)
            self.log_api_call(log_ws, task_id, 'Discord', success)
        
        # Telegram
        if self.config['telegram_bot_token'] and self.config['telegram_chat_id']:
            success = self.api.send_telegram_message(
                self.config['telegram_bot_token'],
                self.config['telegram_chat_id'],
                task_data
            )
            self.log_api_call(log_ws, task_id, 'Telegram', success)
        
        # GitHub
        if self.config['github_token'] and self.config['github_owner'] and self.config['github_repo']:
            url = self.api.create_github_issue(
                self.config['github_token'],
                self.config['github_owner'],
                self.config['github_repo'],
                task_data
            )
            self.log_api_call(log_ws, task_id, 'GitHub', bool(url), url)
    
    def log_api_call(self, ws, task_id, service, success, details=''):
        """Log API integration call"""
        next_row = ws.max_row + 1
        
        log_data = [
            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            task_id,
            service,
            '✅ Success' if success else '❌ Failed',
            'OK' if success else 'Error',
            details
        ]
        
        for idx, value in enumerate(log_data, 1):
            ws.cell(row=next_row, column=idx, value=value)
    
    def save(self):
        """Save the workbook"""
        self.wb.save(self.filename)
        return self.filename

# Create and test
if __name__ == '__main__':
    print("\n🚀 Creating Advanced Task Management System with API Integrations...")
    print("=" * 70)
    
    tms = TaskManagementSystem('task_management_advanced.xlsx')
    
    # Add sample task
    sample_task = {
        'name': 'Implement OAuth2 Authentication',
        'category': 'Development',
        'assigned_to': 'Alice Developer',
        'priority': 'High',
        'due_date': (datetime.now() + timedelta(days=7)).strftime('%Y-%m-%d'),
        'status': 'Not Started',
        'progress': 0,
        'description': 'Add OAuth2 authentication to the API with JWT tokens'
    }
    
    print("\n📝 Adding sample task...")
    task_id = tms.add_task(sample_task, notify=False)
    
    # Get weather-based deadline suggestion
    print("\n🌤️ Getting weather-based deadline suggestion...")
    weather_suggestion = tms.api.suggest_deadline_based_on_weather()
    if weather_suggestion:
        print(f"   Suggested deadline: {weather_suggestion['suggested_date']}")
        print(f"   Reason: {weather_suggestion['reason']}")
    
    print("\n" + "=" * 70)
    print("✅ ADVANCED SYSTEM CREATED!")
    print("=" * 70)
    print(f"\n📁 File: task_management_advanced.xlsx")
    print("\n🔌 Available API Integrations:")
    print("   • Slack notifications")
    print("   • Discord webhooks")
    print("   • Telegram bot messages")
    print("   • GitHub issue creation")
    print("   • Trello card creation")
    print("   • Notion page creation")
    print("   • Airtable records")
    print("   • Weather-based deadline suggestions")
    print("   • Productivity tips")
    print("   • Timezone conversions")
    print("\n💡 To enable notifications, set environment variables:")
    print("   export SLACK_WEBHOOK_URL='your-webhook-url'")
    print("   export DISCORD_WEBHOOK_URL='your-webhook-url'")
    print("   export TELEGRAM_BOT_TOKEN='your-token'")
    print("   export TELEGRAM_CHAT_ID='your-chat-id'")
    print("=" * 70 + "\n")

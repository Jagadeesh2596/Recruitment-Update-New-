import json
import sys
from datetime import datetime
from recruitment_agent import RecruitmentAgent
import sqlite3
import os

class WebIntegrationAgent:
    def __init__(self):
        self.db_path = 'recruitment_web.db'
        self.init_database()
        
    def init_database(self):
        """Initialize SQLite database for web integration"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Admin settings table
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS admin_settings (
            id INTEGER PRIMARY KEY,
            setting_key TEXT UNIQUE,
            setting_value TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        ''')
        
        # Report history table
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS report_history (
            id INTEGER PRIMARY KEY,
            report_data TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            status TEXT
        )
        ''')
        
        # System logs table
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS system_logs (
            id INTEGER PRIMARY KEY,
            log_level TEXT,
            message TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        ''')
        
        # Insert default settings with ASCII-safe template
        default_settings = [
            ('anthropic_api_key', ''),
            ('email_user', ''),
            ('email_password', ''),
            ('client_emails', '[]'),
            ('schedule_frequency', 'weekly'),
            ('schedule_day', 'tuesday'),
            ('schedule_time', '09:00'),
            ('claude_model', 'claude-3-sonnet-20240229'),
            ('system_prompt', 'Analyze this pharmaceutical survey recruitment status and provide professional insights.'),
            ('email_template', '''Subject: Weekly Recruitment Update - {project_name}

Dear Valued Client,

Weekly recruitment progress update for {project_name}:

RECRUITMENT SUMMARY:
================================================================

Total Target: {total_quota} respondents
Current Completes: {overall_completes} respondents
Completion Rate: {completion_percentage}%

SEGMENT BREAKDOWN:
{segments_summary}

AI ANALYSIS:
{analysis}

Report generated: {analysis_date}

Best regards,
Survey Operations Team''')
        ]
        
        for key, value in default_settings:
            cursor.execute('''
            INSERT OR IGNORE INTO admin_settings (setting_key, setting_value)
            VALUES (?, ?)
            ''', (key, value))
        
        conn.commit()
        conn.close()
    
    def get_setting(self, key):
        """Get a setting from database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT setting_value FROM admin_settings WHERE setting_key = ?', (key,))
        result = cursor.fetchone()
        conn.close()
        return result[0] if result else None
    
    def update_setting(self, key, value):
        """Update a setting in database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
        INSERT OR REPLACE INTO admin_settings (setting_key, setting_value, updated_at)
        VALUES (?, ?, ?)
        ''', (key, value, datetime.now()))
        conn.commit()
        conn.close()
    
    def log_message(self, level, message):
        """Log a message to database"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            safe_message = str(message).encode('ascii', 'ignore').decode('ascii')
            cursor.execute('''
            INSERT INTO system_logs (log_level, message)
            VALUES (?, ?)
            ''', (level, safe_message))
            conn.commit()
            conn.close()
        except Exception as e:
            print(f"Error logging message: {e}")
    
    def generate_report_for_web(self):
        try:
            self.log_message('INFO', 'Starting web report generation')
            api_key = self.get_setting('anthropic_api_key')
            if not api_key:
                raise ValueError("Anthropic API key not configured")
            
            agent = RecruitmentAgent(api_key)
            excel_file = agent.fetch_excel_data()
            if not excel_file:
                raise ValueError("Failed to fetch Excel data")
            
            project_data = agent.process_client_summary(excel_file)
            if not project_data:
                raise ValueError("Failed to process client summary")
            
            system_prompt = self.get_setting('system_prompt')
            custom_analysis = self.get_claude_analysis(project_data, system_prompt, api_key)
            report = self.generate_custom_report(project_data, custom_analysis)
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
            INSERT INTO report_history (report_data, status)
            VALUES (?, ?)
            ''', (json.dumps({
                'project_data': project_data,
                'analysis': custom_analysis,
                'report': report
            }, ensure_ascii=True), 'success'))
            conn.commit()
            conn.close()
            
            self.log_message('INFO', 'Report generated successfully')
            
            return {
                'success': True,
                'project_data': project_data,
                'analysis': custom_analysis,
                'report': report,
                'timestamp': datetime.now().isoformat()
            }
            
        except Exception as e:
            error_msg = f"Error generating report: {str(e)}"
            self.log_message('ERROR', error_msg)
            return {
                'success': False,
                'error': error_msg,
                'timestamp': datetime.now().isoformat()
            }

    def get_claude_analysis(self, project_data, system_prompt, api_key):
        try:
            import anthropic
            claude = anthropic.Anthropic(api_key=api_key)
            
            segments_text = ""
            for segment, data in project_data['segments'].items():
                segments_text += f"\n{segment}:\n"
                for category, value in data.items():
                    segments_text += f"  - {category}: {value}\n"
            
            prompt = f"""
            {system_prompt}
            
            PROJECT: {project_data['project_name']}
            Total Target: {project_data['total_quota']} respondents
            Current Completes: {project_data['overall_completes']} respondents
            Completion Rate: {project_data['completion_percentage']*100:.0f}%
            
            SEGMENTS:{segments_text}
            
            Provide:
            1. Status assessment (On Track/Behind/Ahead)
            2. Key insights
            3. Professional summary for client
            """
            
            claude_model = self.get_setting('claude_model') or 'claude-3-sonnet-20240229'
            response = claude.messages.create(
                model=claude_model,
                max_tokens=500,
                messages=[{"role": "user", "content": prompt}]
            )
            
            analysis = response.content[0].text
            return analysis.encode('ascii', 'ignore').decode('ascii')
        except Exception as e:
            return f"Manual Analysis: Project at {project_data['completion_percentage']*100:.0f}% completion rate."

    def generate_custom_report(self, project_data, analysis):
        template = self.get_setting('email_template')
        
        segments_summary = ""
        for segment, data in project_data['segments'].items():
            segments_summary += f"\n{segment}:\n"
            for category, value in data.items():
                segments_summary += f"  * {category}: {value} completes\n"
        
        report = template.format(
            project_name=project_data['project_name'],
            total_quota=project_data['total_quota'],
            overall_completes=project_data['overall_completes'],
            completion_percentage=project_data['completion_percentage']*100,
            segments_summary=segments_summary,
            analysis=analysis,
            analysis_date=project_data['analysis_date']
        )
        return report.encode('ascii', 'ignore').decode('ascii')
    
    def send_emails_to_clients(self):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
            SELECT report_data FROM report_history 
            WHERE status = 'success' 
            ORDER BY created_at DESC LIMIT 1
            ''')
            result = cursor.fetchone()
            conn.close()
            
            if not result:
                raise ValueError("No recent report found")
            
            report_data = json.loads(result[0])
            report = report_data['report']
            
            email_user = self.get_setting('email_user')
            email_password = self.get_setting('email_password')
            client_emails = json.loads(self.get_setting('client_emails') or '[]')
            api_key = self.get_setting('anthropic_api_key')
            
            if not all([email_user, email_password, api_key]):
                raise ValueError("Email credentials not configured")
            
            agent = RecruitmentAgent(api_key)
            success_count = 0
            
            for client_email in client_emails:
                success = agent.send_email(client_email, report, email_user, email_password)
                if success:
                    success_count += 1
                    self.log_message('INFO', f'Email sent successfully to {client_email}')
                else:
                    self.log_message('ERROR', f'Failed to send email to {client_email}')
            
            return {
                'success': True,
                'sent_count': success_count,
                'total_clients': len(client_emails)
            }
            
        except Exception as e:
            error_msg = f"Error sending emails: {str(e)}"
            self.log_message('ERROR', error_msg)
            return {
                'success': False,
                'error': error_msg
            }

if __name__ == "__main__":
    import sys
    if sys.platform == "win32":
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.detach())
    
    web_agent = WebIntegrationAgent()
    
    if len(sys.argv) > 1:
        command = sys.argv[1]
        
        if command == "generate_report":
            result = web_agent.generate_report_for_web()
            print(json.dumps(result, ensure_ascii=True))
        
        elif command == "send_emails":
            result = web_agent.send_emails_to_clients()
            print(json.dumps(result, ensure_ascii=True))
        
        elif command == "get_setting":
            if len(sys.argv) > 2:
                setting = web_agent.get_setting(sys.argv[2])
                print(json.dumps({'value': setting}, ensure_ascii=True))
        
        elif command == "update_setting":
            if len(sys.argv) > 3:
                web_agent.update_setting(sys.argv[2], sys.argv[3])
                print(json.dumps({'success': True}, ensure_ascii=True))

        elif command == "init_db":
            web_agent.init_database()
            print(json.dumps({"success": True, "message": "Database initialized"}, ensure_ascii=True))
    
    else:
        print("Usage: python web_integration.py [generate_report|send_emails|get_setting|update_setting|init_db]")

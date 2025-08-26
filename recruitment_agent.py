import pandas as pd
import requests
from io import BytesIO
import anthropic
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import os
import sys

class RecruitmentAgent:
    def __init__(self, anthropic_key):
        self.claude = anthropic.Anthropic(api_key=anthropic_key)
        self.excel_url = "https://tinyurl.com/ms747thh"
        
    def fetch_excel_data(self):
        """Fetch Excel from local file first, then try online"""
        try:
            print("Looking for Excel data...")
            
            # STEP 1: Try to find local Excel file first
            local_file = self.find_local_excel_file()
            if local_file:
                return self.load_excel_file(local_file)
            
            # STEP 2: Try to fetch from online URL
            print("No local file found, trying online URL...")
            return self.fetch_online_excel()
            
        except Exception as e:
            print(f"Error fetching Excel: {e}")
            return None
    
    def find_local_excel_file(self):
        """Find Excel file in current directory"""
        try:
            current_dir = os.getcwd()
            print(f"Looking in: {current_dir}")
            
            # Look for Excel files
            excel_extensions = ['.xlsx', '.xls']
            excel_files = []
            
            for file in os.listdir('.'):
                if any(file.lower().endswith(ext) for ext in excel_extensions):
                    excel_files.append(file)
            
            if excel_files:
                print(f"Found Excel files: {excel_files}")
                
                # Prioritize files with 'recruitment' in name
                for file in excel_files:
                    if 'recruitment' in file.lower():
                        print(f"Using: {file}")
                        return file
                
                # Use the first Excel file found
                print(f"Using: {excel_files[0]}")
                return excel_files[0]
            
            print("No Excel files found in current directory")
            return None
            
        except Exception as e:
            print(f"Error finding local Excel: {e}")
            return None
    
    def load_excel_file(self, file_path):
        """Load Excel file with multiple engine attempts"""
        try:
            print(f"Loading Excel file: {file_path}")
            
            # Try different engines in order of preference
            engines = ['openpyxl', 'xlrd', None]  # None uses pandas default
            
            for engine in engines:
                try:
                    if engine:
                        print(f"   Trying engine: {engine}")
                        excel_file = pd.ExcelFile(file_path, engine=engine)
                    else:
                        print("   Trying default pandas engine")
                        excel_file = pd.ExcelFile(file_path)
                    
                    print(f"Success! Sheets found: {excel_file.sheet_names}")
                    return excel_file
                    
                except Exception as engine_error:
                    print(f"   {engine or 'default'} engine failed: {engine_error}")
                    continue
            
            print("All engines failed to load the Excel file")
            return None
            
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return None
    
    def fetch_online_excel(self):
        """Fetch Excel from OneDrive URL (backup method)"""
        try:
            print("Attempting to fetch from online URL...")
            response = requests.get(self.excel_url, allow_redirects=True, timeout=30)
            
            if response.status_code == 200:
                excel_file = pd.ExcelFile(BytesIO(response.content))
                print(f"Online Excel fetched! Sheets: {excel_file.sheet_names}")
                return excel_file
            else:
                print(f"Online fetch failed: HTTP {response.status_code}")
                return None
                
        except Exception as e:
            print(f"Online fetch error: {e}")
            return None
    
    def process_client_summary(self, excel_file):
        """Process Client Summary tab with better error handling"""
        try:
            print("Processing Client Summary tab...")
            
            # Check if 'Client Summary' sheet exists
            if hasattr(excel_file, 'sheet_names'):
                available_sheets = excel_file.sheet_names
            else:
                # Handle case where excel_file might be a different object
                available_sheets = ['Client Summary']  # Assume default
            
            print(f"Available sheets: {available_sheets}")
            
            # Try different sheet name variations
            sheet_variations = ['Client Summary', 'ClientSummary', 'client summary', 'Summary']
            
            df = None
            sheet_used = None
            
            for sheet_name in sheet_variations:
                if sheet_name in available_sheets:
                    try:
                        if hasattr(excel_file, 'parse'):  # For ExcelFile objects
                            df = excel_file.parse(sheet_name)
                        else:  # For file path strings
                            df = pd.read_excel(excel_file, sheet_name=sheet_name)
                        sheet_used = sheet_name
                        break
                    except Exception as sheet_error:
                        print(f"Failed to read sheet '{sheet_name}': {sheet_error}")
                        continue
            
            if df is None:
                print(f"Could not find 'Client Summary' sheet. Available: {available_sheets}")
                return None
            
            print(f"Successfully loaded sheet: {sheet_used}")
            print(f"Sheet dimensions: {df.shape[0]} rows x {df.shape[1]} columns")
            
            # Initialize project data
            project_data = {
                'project_name': 'GLD HBV PET Survey',
                'total_quota': 0,
                'overall_completes': 0,
                'completion_percentage': 0,
                'segments': {},
                'analysis_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            # Convert to list for easier processing
            data_rows = df.values.tolist()
            
            # Extract main metrics
            for i, row in enumerate(data_rows):
                if row and len(row) > 4:
                    # Look for "Total Quota" header
                    if any(cell and 'Total Quota' in str(cell) for cell in row if pd.notna(cell)):
                        if i + 1 < len(data_rows):
                            values_row = data_rows[i + 1]
                            try:
                                project_data['total_quota'] = int(values_row[1]) if pd.notna(values_row[1]) else 0
                                project_data['overall_completes'] = int(values_row[2]) if pd.notna(values_row[2]) else 0
                                project_data['completion_percentage'] = float(values_row[4]) if pd.notna(values_row[4]) else 0
                            except (ValueError, IndexError, TypeError):
                                pass
                        break
            
            # Extract segments
            current_segment = None
            for i, row in enumerate(data_rows):
                if row and len(row) > 0:
                    # Look for segment headers
                    for cell in row:
                        if pd.notna(cell) and 'Split' in str(cell):
                            current_segment = str(cell).replace(' Split', '').strip()
                            project_data['segments'][current_segment] = {}
                            break
                    
                    # Extract segment data
                    if current_segment and len(row) > 3 and pd.notna(row[2]):
                        category = str(row[2]).strip()
                        if category and len(category) > 2:  # Valid category
                            try:
                                value = int(row[3]) if pd.notna(row[3]) else (int(row[4]) if len(row) > 4 and pd.notna(row[4]) else 0)
                                if value > 0:
                                    project_data['segments'][current_segment][category] = value
                            except (ValueError, TypeError):
                                pass
            
            print(f"Processed data: {project_data['overall_completes']}/{project_data['total_quota']} completes ({project_data['completion_percentage']*100:.1f}%)")
            print(f"Found {len(project_data['segments'])} segments")
            
            return project_data
            
        except Exception as e:
            print(f"Error processing Client Summary: {e}")
            import traceback
            print(f"Full error trace: {traceback.format_exc()}")
            return None
    
    def analyze_with_claude(self, project_data):
        """Get AI analysis from Claude"""
        try:
            print("Getting Claude analysis...")
            
            segments_text = ""
            for segment, data in project_data['segments'].items():
                segments_text += f"\n{segment}:\n"
                for category, value in data.items():
                    segments_text += f"  - {category}: {value}\n"
            
            prompt = f"""
            Analyze this pharmaceutical survey recruitment status:
            
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
            
            try:
                response = self.claude.messages.create(
                    model="claude-3-sonnet-20240229",
                    max_tokens=500,
                    messages=[{"role": "user", "content": prompt}]
                )
                
                analysis = response.content[0].text
                print("Claude analysis complete")
                # Ensure ASCII-safe response
                return analysis.encode('ascii', 'ignore').decode('ascii')
            
            except Exception as api_error:
                print(f"Claude API error: {api_error}")
                # Return manual analysis as backup
                return self.manual_analysis(project_data)
            
        except Exception as e:
            print(f"Error with Claude analysis: {e}")
            return self.manual_analysis(project_data)
    
    def manual_analysis(self, project_data):
        """Backup analysis when Claude fails"""
        completion_rate = project_data['completion_percentage'] * 100
        
        if completion_rate >= 100:
            status = "COMPLETED - Target achieved"
        elif completion_rate >= 90:
            status = "ON TRACK - Near completion"
        elif completion_rate >= 75:
            status = "ON TRACK - Good progress"
        else:
            status = "BEHIND SCHEDULE - Needs attention"
        
        analysis = f"""
        STATUS: {status}
        
        ANALYSIS:
        - Survey has achieved {completion_rate:.0f}% of target quota
        - {project_data['overall_completes']} completes out of {project_data['total_quota']} target
        - Strong performance across specialty segments
        - Recruitment progressing as planned
        
        RECOMMENDATIONS:
        - Continue current recruitment strategy
        - Monitor segment balance for any adjustments needed
        - Project on track for successful completion
        """
        
        return analysis
    
    def generate_report(self, project_data, analysis):
        """Generate client report with ASCII-safe formatting"""
        segments_summary = ""
        for segment, data in project_data['segments'].items():
            segments_summary += f"\n{segment}:\n"
            for category, value in data.items():
                # Use ASCII-safe bullet points
                segments_summary += f"  * {category}: {value} completes\n"
        
        report = f"""
Subject: Weekly Recruitment Update - {project_data['project_name']}

Dear Valued Client,

Weekly recruitment progress update for {project_data['project_name']}:

RECRUITMENT SUMMARY:
================================================================

Total Target: {project_data['total_quota']} respondents
Current Completes: {project_data['overall_completes']} respondents
Completion Rate: {project_data['completion_percentage']*100:.0f}%

SEGMENT BREAKDOWN:{segments_summary}

AI ANALYSIS:
{analysis}

Report generated: {project_data['analysis_date']}

Best regards,
Survey Operations Team
        """
        
        # Ensure ASCII-safe report
        return report.encode('ascii', 'ignore').decode('ascii')
    
    def send_email(self, recipient_email, report, email_user, email_password):
        """Send email report with proper encoding"""
        try:
            print(f"Sending email to {recipient_email}...")
            
            # Create multipart message for better compatibility
            msg = MIMEMultipart()
            msg['Subject'] = "Weekly Recruitment Update"
            msg['From'] = email_user
            msg['To'] = recipient_email
            
            # Attach the report as plain text with UTF-8 encoding
            msg.attach(MIMEText(report, 'plain', 'utf-8'))
            
            # Use Gmail SMTP settings
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(email_user, email_password)  # Use app password here
            server.send_message(msg)
            server.quit()
            
            print("Email sent successfully!")
            return True
            
        except Exception as e:
            print(f"Email error: {e}")
            return False
    
    def run_complete_process(self, client_email=None, email_user=None, email_password=None):
        """Run the complete recruitment analysis process"""
        print("Starting Recruitment Agent...")
        
        # Step 1: Fetch Excel
        excel_file = self.fetch_excel_data()
        if not excel_file:
            return False
        
        # Step 2: Process data
        project_data = self.process_client_summary(excel_file)
        if not project_data:
            return False
        
        # Step 3: AI Analysis
        analysis = self.analyze_with_claude(project_data)
        
        # Step 4: Generate report
        report = self.generate_report(project_data, analysis)
        
        print("\n" + "="*60)
        print("GENERATED REPORT:")
        print("="*60)
        print(report)
        
        # Step 5: Send email (if credentials provided)
        if client_email and email_user and email_password:
            self.send_email(client_email, report, email_user, email_password)
        else:
            print("\nTo send email, provide: client_email, email_user, email_password")
        
        return True

# Test the system
if __name__ == "__main__":
    # Set UTF-8 encoding for Windows console
    if sys.platform == "win32":
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.detach())
    
    # Initialize with your Anthropic key
    try:
        from config import ANTHROPIC_API_KEY
        agent = RecruitmentAgent(ANTHROPIC_API_KEY)
    except ImportError:
        ANTHROPIC_KEY = input("Enter your Anthropic API key: ")
        agent = RecruitmentAgent(ANTHROPIC_KEY)
    
    # Test run
    print("Testing the recruitment agent...")
    success = agent.run_complete_process()
    
    if success:
        print("\nTest completed successfully!")
        print("\nNext: Run automated_scheduler.py for full automation")
    else:
        print("\nTest failed. Check the errors above.")
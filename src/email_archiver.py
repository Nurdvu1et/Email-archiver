import imaplib
import email
import os
import datetime
import sqlite3
import logging
from email.header import decode_header
import time
from typing import Optional, List, Dict, Any

class EmailArchiver:
    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.db_conn = sqlite3.connect('email_archive.db')
        self.max_emails_per_session = 100
        self.error_count = 0
        self.max_errors = 5
        self.setup_logging()
        self.create_database()
        
    def setup_logging(self):
        """Configure detailed logging"""
        logging.basicConfig(
            filename='email_archiver.log',
            level=logging.DEBUG if self.config.get('debug') else logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            filemode='a'  # Append mode
        )
        self.logger = logging.getLogger(__name__)
        console = logging.StreamHandler()
        console.setLevel(logging.INFO)
        self.logger.addHandler(console)

    def create_database(self):
        """Initialize database with error handling"""
        try:
            cursor = self.db_conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS archived_emails (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email_id TEXT UNIQUE,
                    subject TEXT,
                    sender TEXT,
                    date_received TEXT,
                    attachments TEXT,
                    keywords TEXT,
                    archive_path TEXT,
                    processed_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS search_index (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email_id INTEGER,
                    keyword TEXT,
                    FOREIGN KEY(email_id) REFERENCES archived_emails(id)
                )
            ''')
            self.db_conn.commit()
        except Exception as e:
            self.logger.error(f"Database setup failed: {str(e)}")
            raise

    def connect_to_mailbox(self) -> Optional[imaplib.IMAP4_SSL]:
        """Enhanced IMAP connection with debugging"""
        try:
            self.logger.debug(f"Connecting to {self.config['imap_server']}")
            mail = imaplib.IMAP4_SSL(self.config['imap_server'])
            mail.debug = 4 if self.config.get('debug') else 0
            self.logger.info(f"Logging in as {self.config['email']}")
            mail.login(self.config['email'], self.config['password'])
            mail.select('inbox')
            return mail
        except Exception as e:
            self.logger.error(f"IMAP connection failed: {str(e)}")
            return None

    def safe_decode_header(self, header: Optional[str]) -> str:
        """More robust header decoding"""
        if header is None:
            return "(No Header)"
        try:
            decoded, encoding = decode_header(header)[0]
            if isinstance(decoded, bytes):
                return decoded.decode(encoding or 'utf-8', errors='replace')
            return str(decoded)
        except Exception as e:
            self.logger.warning(f"Header decode failed: {str(e)}")
            return str(header)[:100]

    def process_emails(self):
        """Main email processing method with error handling"""
        mail = self.connect_to_mailbox()
        if not mail:
            self.logger.error("Cannot proceed without IMAP connection")
            return

        try:
            # Only process unseen emails to prevent duplicates
            status, messages = mail.search(None, '(UNSEEN)')
            if status != 'OK':
                self.logger.warning("Email search failed or no new emails")
                return

            email_ids = messages[0].split()
            self.logger.info(f"Found {len(email_ids)} new emails to process")

            processed_count = 0
            for email_id in email_ids[:self.max_emails_per_session]:
                if self.error_count >= self.max_errors:
                    self.logger.error("Reached maximum error limit, stopping")
                    break
                
                if self.process_single_email(mail, email_id):
                    processed_count += 1

            self.logger.info(f"Successfully processed {processed_count} emails")
            
        except Exception as e:
            self.logger.error(f"Fatal error: {str(e)}", exc_info=True)
        finally:
            try:
                mail.close()
                mail.logout()
            except:
                pass

    def process_single_email(self, mail: imaplib.IMAP4_SSL, email_id: bytes) -> bool:
        """Enhanced single email processing with debug output"""
        try:
            status, msg_data = mail.fetch(email_id, '(RFC822)')
            if status != 'OK':
                self.logger.warning(f"Fetch failed for email {email_id}")
                return False

            raw_email = msg_data[0][1]
            email_message = email.message_from_bytes(raw_email)
            
            # Debug email structure
            self.logger.debug("Email structure:")
            for header in ['Subject', 'From', 'Date', 'Message-ID']:
                self.logger.debug(f"{header}: {email_message[header]}")

            subject = self.safe_decode_header(email_message['Subject'])
            sender = self.safe_decode_header(email_message['From'])
            date_received = self.safe_decode_header(email_message['Date'])
            
            # Create archive directory
            archive_date = datetime.datetime.now().strftime('%Y-%m-%d')
            safe_sender = "".join(
                c if c.isalnum() or c in ('@', '.', '-', '_') else '_' 
                for c in sender.split('@')[0]
            )[:50]
            
            archive_dir = os.path.join(
                self.config['archive_root'],
                archive_date,
                safe_sender,
                email_id.decode()[:10]  # Add email ID to path for uniqueness
            )
            os.makedirs(archive_dir, exist_ok=True)
            
            # Process attachments with detailed logging
            attachments = []
            for part_num, part in enumerate(email_message.walk()):
                content_type = part.get_content_type()
                disposition = str(part.get('Content-Disposition'))
                
                self.logger.debug(f"Part {part_num}: {content_type} | {disposition}")
                
                if part.is_multipart():
                    continue
                    
                if 'attachment' not in disposition.lower():
                    self.logger.debug("Skipping non-attachment part")
                    continue

                filename = part.get_filename()
                if not filename:
                    ext = content_type.split('/')[-1]
                    filename = f"attachment_{len(attachments)}.{ext}"
                
                filename = self.safe_decode_header(filename)
                safe_filename = "".join(
                    c if c.isalnum() or c in ('.', '-', '_') else '_'
                    for c in filename
                )[:100]
                
                filepath = os.path.join(archive_dir, safe_filename)
                try:
                    payload = part.get_payload(decode=True)
                    if payload:
                        with open(filepath, 'wb') as f:
                            f.write(payload)
                        attachments.append(safe_filename)
                        self.logger.info(f"Saved attachment: {safe_filename}")
                    else:
                        self.logger.warning("Empty attachment payload")
                except Exception as e:
                    self.logger.error(f"Failed to save {filename}: {str(e)}")
                    continue
                
            if attachments:
                self.store_email_metadata(
                    email_id.decode(),
                    subject,
                    sender,
                    date_received,
                    attachments,
                    archive_dir
                )
                
                if self.config.get('delete_after_archive', False):
                    mail.store(email_id, '+FLAGS', '\\Deleted')
                    self.logger.debug("Marked email for deletion")
                
                return True
            else:
                self.logger.warning("No attachments found in email")
                return False
            
        except Exception as e:
            self.error_count += 1
            self.logger.error(f"Error processing email: {str(e)}", exc_info=True)
            return False

    def store_email_metadata(self, email_id: str, subject: str, sender: str, 
                           date_received: str, attachments: List[str], archive_path: str):
        """Store email metadata in database with search indexing"""
        try:
            cursor = self.db_conn.cursor()
            cursor.execute('''
                INSERT INTO archived_emails 
                (email_id, subject, sender, date_received, attachments, archive_path)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (email_id, subject, sender, date_received, ', '.join(attachments), archive_path))
            
            # Improved keyword extraction
            keywords = set()
            for text in [subject, sender]:
                for word in text.lower().split():
                    clean_word = ''.join(c for c in word if c.isalnum())
                    if 3 <= len(clean_word) <= 30:  # Reasonable word length
                        keywords.add(clean_word)
            
            for keyword in keywords:
                cursor.execute('''
                    INSERT INTO search_index (email_id, keyword)
                    VALUES (?, ?)
                ''', (cursor.lastrowid, keyword))
            
            self.db_conn.commit()
        except Exception as e:
            self.db_conn.rollback()
            self.logger.error(f"Database error: {str(e)}")

    def search_emails(self, query: str) -> List[Dict[str, Any]]:
        """Search archived emails with enhanced capabilities"""

        print("DEBUG: Checking database connection...")
        print("DEBUG: Database path:", os.path.abspath('email_archive.db'))
        print("DEBUG: Tables exist:", self.db_conn.execute("SELECT name FROM sqlite_master").fetchall())
        try:
            if not hasattr(self, 'db_conn'):
                self.logger.error("Database connection not established")
                return []

            cursor = self.db_conn.cursor()
            query = query.strip().lower()
            
            # Search across multiple fields
            cursor.execute('''
                SELECT DISTINCT e.subject, e.sender, e.date_received, 
                       e.attachments, e.archive_path
                FROM archived_emails e
                LEFT JOIN search_index s ON e.id = s.email_id
                WHERE e.subject LIKE ? OR 
                      e.sender LIKE ? OR 
                      e.attachments LIKE ? OR
                      e.date_received LIKE ? OR
                      s.keyword LIKE ?
                ORDER BY e.date_received DESC
                LIMIT 100
            ''', (f'%{query}%', f'%{query}%', f'%{query}%', f'%{query}%', f'%{query}%'))
            
            results = [{
                'subject': row[0],
                'sender': row[1],
                'date': row[2],
                'attachments': row[3],
                'path': row[4]
            } for row in cursor.fetchall()]

            self.logger.debug(f"Search for '{query}' returned {len(results)} results")
            return results
            
        except Exception as e:
            self.logger.error(f"Search failed: {str(e)}")
            return []

    def cleanup_mailbox(self):
        """Permanently delete marked emails with confirmation"""
        mail = self.connect_to_mailbox()
        if not mail:
            return
            
        try:
            mail.select('inbox')
            response = input("Permanently delete all marked emails? (y/n): ")
            if response.lower() == 'y':
                mail.expunge()
                self.logger.info("Mailbox cleanup completed")
            else:
                self.logger.info("Cleanup canceled by user")
        except Exception as e:
            self.logger.error(f"Cleanup failed: {str(e)}")
        finally:
            mail.close()
            mail.logout()

def load_config() -> Dict[str, Any]:
    """Load configuration with debug options"""
    return {
        'imap_server': os.getenv('IMAP_SERVER', 'imap.gmail.com'),
        'email': os.getenv('EMAIL'),
        'password': os.getenv('PASSWORD'),
        'archive_root': os.getenv('ARCHIVE_ROOT', './email_archive'),
        'delete_after_archive': os.getenv('DELETE_AFTER_ARCHIVE', 'False').lower() == 'true',
    }

def main_menu(archiver: EmailArchiver):
    while True:
        print("\nEmail Archiver Menu:")
        print("1. Process and archive new emails")
        print("2. Search archived emails")
        print("3. Cleanup mailbox")
        print("4. Exit")
        
        try:
            choice = input("Enter your choice (1-4): ").strip()
            
            if choice == '1':
                print("Processing emails...")
                archiver.process_emails()
            elif choice == '2':
                query = input("Enter search term: ").strip()
                results = archiver.search_emails(query)  
                print(f"\nFound {len(results)} results:")
                for i, result in enumerate(results, 1):
                    print(f"\nResult {i}:")
                    print(f"Subject: {result['subject']}")
                    print(f"From: {result['sender']}")
                    print(f"Date: {result['date']}")
                    print(f"Attachments: {result['attachments']}")
                    print(f"Path: {result['path']}")
            elif choice == '3':
                archiver.cleanup_mailbox()
            elif choice == '4':
                print("Exiting...")
                break
            else:
                print("Invalid choice, please enter 1-4")
            
        except KeyboardInterrupt:
            print("\nOperation canceled")
            break
        except Exception as e:
            print(f"Error: {str(e)}")

if __name__ == "__main__":
    try:
        config = load_config()
        if not all([config['email'], config['password']]):
            raise ValueError("Email and password must be configured")
            
        print(f"\nStarting Email Archiver")
        print(f"IMAP Server: {config['imap_server']}")
        print(f"Archive Location: {os.path.abspath(config['archive_root'])}")
        
        archiver = EmailArchiver(config)
        main_menu(archiver)
    except Exception as e:
        print(f"Failed to start: {str(e)}")
        logging.error(f"Startup failed: {str(e)}", exc_info=True)
import sqlite3
conn = sqlite3.connect('email_archive.db')
cursor = conn.cursor()
cursor.execute("SELECT COUNT(*) FROM archived_emails")
print(f"Total archived emails: {cursor.fetchone()[0]}")
cursor.execute("SELECT * FROM archived_emails LIMIT 1")
print("Sample email:", cursor.fetchone())
conn.close()
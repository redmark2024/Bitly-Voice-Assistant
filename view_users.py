from database import Database
import mysql.connector
from datetime import datetime

def view_all_users():
    try:
        # Create database connection
        conn = mysql.connector.connect(
            host='localhost',
            user='root',
            password='',
            database='bitly_assistant',
            port=3306,
            auth_plugin='mysql_native_password'
        )
        
        cursor = conn.cursor()
        
        # Get all users
        cursor.execute("""
            SELECT id, username, email, created_at 
            FROM users 
            ORDER BY created_at DESC
        """)
        
        users = cursor.fetchall()
        
        print("\n=== Registered Users in Database ===")
        print("Total users:", len(users))
        print("\n" + "="*80)
        
        for user in users:
            print(f"\n👤 User ID: {user[0]}")
            print(f"📝 Username: {user[1]}")
            print(f"📧 Email: {user[2]}")
            print(f"🕒 Created: {user[3]}")
            print("-"*80)
        
        cursor.close()
        conn.close()
        
    except mysql.connector.Error as err:
        print(f"Error: {err}")

if __name__ == "__main__":
    print("Fetching user data from database...")
    view_all_users() 
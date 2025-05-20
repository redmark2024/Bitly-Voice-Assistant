import mysql.connector
from database import Database
import time

def test_connection():
    try:
        print("\n=== Testing Database Connection ===")
        # Create database instance
        db = Database()
        
        # Try to connect to MySQL
        conn = mysql.connector.connect(
            host='localhost',
            user='root',
            password='',
            database='bitly_assistant',
            port=3306,
            auth_plugin='mysql_native_password'
        )
        
        if conn.is_connected():
            print("✅ Successfully connected to MySQL database!")
            print("📊 Database Info:", conn.get_server_info())
            
            # Create cursor and test query
            cursor = conn.cursor()
            cursor.execute("SELECT DATABASE();")
            database = cursor.fetchone()
            print("📁 Connected to database:", database[0])
            
            # Test user registration
            print("\n=== Testing User Registration ===")
            # Generate unique username using timestamp
            timestamp = int(time.time())
            test_username = f"test_user_{timestamp}"
            test_password = "test123"
            test_email = f"test_{timestamp}@example.com"
            
            print(f"Testing with username: {test_username}")
            print(f"Testing with email: {test_email}")
            
            # Try to register a test user
            if db.register_user(test_username, test_password, test_email):
                print("✅ Test user registration successful!")
                
                # Verify the user exists
                if db.verify_user(test_username, test_password):
                    print("✅ Test user verification successful!")
                else:
                    print("❌ Test user verification failed!")
            else:
                print("❌ Test user registration failed!")
            
            # Show registered users
            print("\n=== Registered Users ===")
            cursor.execute("SELECT username, email, created_at FROM users")
            users = cursor.fetchall()
            for user in users:
                print(f"👤 Username: {user[0]}")
                print(f"📧 Email: {user[1]}")
                print(f"🕒 Created: {user[2]}")
                print("---")
            
            # Close connection
            cursor.close()
            conn.close()
            print("\n✅ MySQL connection is closed")
            
            return True
    except mysql.connector.Error as err:
        print("\n❌ Error connecting to MySQL:", err)
        print("\nTroubleshooting steps:")
        print("1. Make sure XAMPP is running")
        print("2. Check if MySQL service is started in XAMPP Control Panel")
        print("3. Verify MySQL is running on port 3306")
        print("4. Check if the root user has no password (default XAMPP setting)")
        print("5. Try accessing http://localhost/phpmyadmin in your browser")
        return False

if __name__ == "__main__":
    print("Starting database connection test...")
    test_connection() 
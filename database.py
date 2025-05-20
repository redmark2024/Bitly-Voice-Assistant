import mysql.connector
import hashlib
import os

class Database:
    def __init__(self):
        # MySQL connection configuration for XAMPP phpMyAdmin
        self.config = {
            'host': 'localhost',
            'user': 'root',
            'password': '',  # Default XAMPP MySQL password is empty
            'database': 'bitly_assistant',
            'port': 3306,  # Default MySQL port
            'auth_plugin': 'mysql_native_password'  # Add this for XAMPP compatibility
        }
        try:
            self.create_database()
            self.create_tables()
            print("Database initialization completed successfully!")
        except Exception as e:
            print(f"Error during database initialization: {e}")
            raise
    
    def create_database(self):
        """Create the database if it doesn't exist"""
        try:
            # Connect to MySQL server without database
            conn = mysql.connector.connect(
                host=self.config['host'],
                user=self.config['user'],
                password=self.config['password'],
                port=self.config['port'],
                auth_plugin=self.config['auth_plugin']
            )
            cursor = conn.cursor()
            
            # Create database if it doesn't exist
            cursor.execute(f"CREATE DATABASE IF NOT EXISTS {self.config['database']}")
            print(f"Database '{self.config['database']}' created or already exists")
            
            conn.close()
        except Exception as e:
            print(f"Error creating database: {e}")
            raise
    
    def create_tables(self):
        """Create necessary tables in the database"""
        try:
            conn = mysql.connector.connect(**self.config)
            cursor = conn.cursor()
            
            # Create users table with additional fields
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS users (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    username VARCHAR(50) UNIQUE NOT NULL,
                    password VARCHAR(64) NOT NULL,
                    email VARCHAR(100) UNIQUE NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    last_login TIMESTAMP NULL,
                    is_active BOOLEAN DEFAULT TRUE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
            ''')
            
            print("Users table created or already exists")
            conn.commit()
            conn.close()
        except Exception as e:
            print(f"Error creating tables: {e}")
            raise
    
    def hash_password(self, password):
        """Hash a password using SHA-256"""
        return hashlib.sha256(password.encode()).hexdigest()
    
    def register_user(self, username, password, email):
        """Register a new user"""
        try:
            conn = mysql.connector.connect(**self.config)
            cursor = conn.cursor()
            
            hashed_password = self.hash_password(password)
            
            cursor.execute('''
                INSERT INTO users (username, password, email)
                VALUES (%s, %s, %s)
            ''', (username, hashed_password, email))
            
            conn.commit()
            conn.close()
            return True
        except mysql.connector.IntegrityError:
            return False
        except Exception as e:
            print(f"Error registering user: {e}")
            return False
    
    def verify_user(self, username, password):
        """Verify user credentials"""
        try:
            conn = mysql.connector.connect(**self.config)
            cursor = conn.cursor()
            
            hashed_password = self.hash_password(password)
            
            cursor.execute('''
                SELECT id FROM users
                WHERE username = %s AND password = %s
            ''', (username, hashed_password))
            
            result = cursor.fetchone()
            conn.close()
            
            return result is not None
        except Exception as e:
            print(f"Error verifying user: {e}")
            return False
    
    def user_exists(self, username):
        """Check if a username exists"""
        try:
            conn = mysql.connector.connect(**self.config)
            cursor = conn.cursor()
            
            cursor.execute('SELECT id FROM users WHERE username = %s', (username,))
            result = cursor.fetchone()
            
            conn.close()
            return result is not None
        except Exception as e:
            print(f"Error checking user existence: {e}")
            return False
    
    def email_exists(self, email):
        """Check if an email exists"""
        try:
            conn = mysql.connector.connect(**self.config)
            cursor = conn.cursor()
            
            cursor.execute('SELECT id FROM users WHERE email = %s', (email,))
            result = cursor.fetchone()
            
            conn.close()
            return result is not None
        except Exception as e:
            print(f"Error checking email existence: {e}")
            return False
    
    def reset_password(self, email, new_password):
        """Reset the password for a user by email"""
        try:
            conn = mysql.connector.connect(**self.config)
            cursor = conn.cursor()
            hashed_password = self.hash_password(new_password)
            cursor.execute('''
                UPDATE users SET password = %s WHERE email = %s
            ''', (hashed_password, email))
            conn.commit()
            updated = cursor.rowcount
            conn.close()
            return updated > 0
        except Exception as e:
            print(f"Error resetting password: {e}")
            return False 
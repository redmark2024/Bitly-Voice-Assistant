import customtkinter as ctk
from database import Database
import re

class AuthGUI:
    def __init__(self):
        self.db = Database()
        self.setup_gui()
        
    def setup_gui(self):
        self.window = ctk.CTk()
        self.window.title("Bitly Voice Assistant - Authentication")
        self.window.geometry("400x500")
        
        # Set appearance mode
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # Create main frame
        self.main_frame = ctk.CTkFrame(self.window)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        # Title
        self.title_label = ctk.CTkLabel(
            self.main_frame,
            text="Bitly Voice Assistant",
            font=("Helvetica", 24, "bold")
        )
        self.title_label.pack(pady=20)
        
        # Create login frame
        self.login_frame = ctk.CTkFrame(self.main_frame)
        self.login_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Login widgets
        self.username_entry = ctk.CTkEntry(
            self.login_frame,
            placeholder_text="Username"
        )
        self.username_entry.pack(pady=10, padx=20, fill="x")
        
        self.password_entry = ctk.CTkEntry(
            self.login_frame,
            placeholder_text="Password",
            show="*"
        )
        self.password_entry.pack(pady=10, padx=20, fill="x")
        
        self.login_button = ctk.CTkButton(
            self.login_frame,
            text="Login",
            command=self.login
        )
        self.login_button.pack(pady=10, padx=20, fill="x")
        
        # Switch to signup
        self.switch_to_signup = ctk.CTkButton(
            self.login_frame,
            text="Don't have an account? Sign up",
            command=self.show_signup,
            fg_color="transparent",
            hover_color=("gray70", "gray30")
        )
        self.switch_to_signup.pack(pady=5)
        
        # Forgot Password button
        self.forgot_password_button = ctk.CTkButton(
            self.login_frame,
            text="Forgot Password?",
            command=self.show_reset_password,
            fg_color="transparent",
            hover_color=("gray70", "gray30")
        )
        self.forgot_password_button.pack(pady=5)
        
        # Create signup frame (initially hidden)
        self.signup_frame = ctk.CTkFrame(self.main_frame)
        
        # Signup widgets
        self.signup_username_entry = ctk.CTkEntry(
            self.signup_frame,
            placeholder_text="Username"
        )
        self.signup_username_entry.pack(pady=10, padx=20, fill="x")
        
        self.signup_email_entry = ctk.CTkEntry(
            self.signup_frame,
            placeholder_text="Email"
        )
        self.signup_email_entry.pack(pady=10, padx=20, fill="x")
        
        self.signup_password_entry = ctk.CTkEntry(
            self.signup_frame,
            placeholder_text="Password",
            show="*"
        )
        self.signup_password_entry.pack(pady=10, padx=20, fill="x")
        
        self.signup_confirm_entry = ctk.CTkEntry(
            self.signup_frame,
            placeholder_text="Confirm Password",
            show="*"
        )
        self.signup_confirm_entry.pack(pady=10, padx=20, fill="x")
        
        self.signup_button = ctk.CTkButton(
            self.signup_frame,
            text="Sign Up",
            command=self.signup
        )
        self.signup_button.pack(pady=10, padx=20, fill="x")
        
        # Switch to login
        self.switch_to_login = ctk.CTkButton(
            self.signup_frame,
            text="Already have an account? Login",
            command=self.show_login,
            fg_color="transparent",
            hover_color=("gray70", "gray30")
        )
        self.switch_to_login.pack(pady=5)
        
        # Status label
        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="",
            text_color="red"
        )
        self.status_label.pack(pady=10)
        
    def show_signup(self):
        self.login_frame.pack_forget()
        self.signup_frame.pack(pady=10, padx=20, fill="both", expand=True)
        self.status_label.configure(text="")
        
    def show_login(self):
        self.signup_frame.pack_forget()
        self.login_frame.pack(pady=10, padx=20, fill="both", expand=True)
        self.status_label.configure(text="")
        
    def validate_email(self, email):
        pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
        return re.match(pattern, email) is not None
        
    def login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()
        
        if not username or not password:
            self.status_label.configure(text="Please fill in all fields")
            return
            
        if self.db.verify_user(username, password):
            self.status_label.configure(text="Login successful!", text_color="green")
            self.window.quit()
        else:
            self.status_label.configure(text="Invalid username or password")
            
    def signup(self):
        username = self.signup_username_entry.get()
        email = self.signup_email_entry.get()
        password = self.signup_password_entry.get()
        confirm_password = self.signup_confirm_entry.get()
        
        # Validate input
        if not username or not email or not password or not confirm_password:
            self.status_label.configure(text="Please fill in all fields")
            return
            
        if not self.validate_email(email):
            self.status_label.configure(text="Invalid email format")
            return
            
        if password != confirm_password:
            self.status_label.configure(text="Passwords do not match")
            return
            
        if len(password) < 6:
            self.status_label.configure(text="Password must be at least 6 characters")
            return
            
        if self.db.user_exists(username):
            self.status_label.configure(text="Username already exists")
            return
            
        if self.db.email_exists(email):
            self.status_label.configure(text="Email already registered")
            return
            
        # Register user
        if self.db.register_user(username, password, email):
            self.status_label.configure(text="Registration successful! Please login", text_color="green")
            self.show_login()
        else:
            self.status_label.configure(text="Registration failed")
            
    def show_reset_password(self):
        # Create a new window for password reset
        reset_window = ctk.CTkToplevel(self.window)
        reset_window.title("Reset Password")
        reset_window.geometry("350x250")
        reset_window.grab_set()

        email_label = ctk.CTkLabel(reset_window, text="Enter your registered email:")
        email_label.pack(pady=10)
        email_entry = ctk.CTkEntry(reset_window, placeholder_text="Email")
        email_entry.pack(pady=5, fill="x", padx=20)

        new_password_label = ctk.CTkLabel(reset_window, text="Enter new password:")
        new_password_label.pack(pady=10)
        new_password_entry = ctk.CTkEntry(reset_window, placeholder_text="New Password", show="*")
        new_password_entry.pack(pady=5, fill="x", padx=20)

        confirm_password_label = ctk.CTkLabel(reset_window, text="Confirm new password:")
        confirm_password_label.pack(pady=10)
        confirm_password_entry = ctk.CTkEntry(reset_window, placeholder_text="Confirm Password", show="*")
        confirm_password_entry.pack(pady=5, fill="x", padx=20)

        status_label = ctk.CTkLabel(reset_window, text="", text_color="red")
        status_label.pack(pady=5)

        def reset_action():
            email = email_entry.get()
            new_password = new_password_entry.get()
            confirm_password = confirm_password_entry.get()
            if not email or not new_password or not confirm_password:
                status_label.configure(text="Please fill in all fields")
                return
            if not self.validate_email(email):
                status_label.configure(text="Invalid email format")
                return
            if new_password != confirm_password:
                status_label.configure(text="Passwords do not match")
                return
            if len(new_password) < 6:
                status_label.configure(text="Password must be at least 6 characters")
                return
            if not self.db.email_exists(email):
                status_label.configure(text="Email not registered")
                return
            if self.db.reset_password(email, new_password):
                status_label.configure(text="Password reset successful!", text_color="green")
            else:
                status_label.configure(text="Password reset failed")

        reset_button = ctk.CTkButton(reset_window, text="Reset Password", command=reset_action)
        reset_button.pack(pady=15, fill="x", padx=20)
        
    def start(self):
        self.window.mainloop() 
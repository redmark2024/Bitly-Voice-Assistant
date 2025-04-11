# Bitly-Voice-Assistant
🔊 Voice Control Assistant Windows Application – Hands-Free Control (Without Mouse & Keyboard)


Voice Control Assistant: Features and GUI Details
This is a comprehensive Python-based Voice Control Assistant application that provides voice-activated computer control with a modern GUI. Here's a breakdown of its features and interface details:
Core Features
Voice Recognition and Control

Uses speech recognition to process voice commands
Wake word system ("assistant" by default) for activation
Text-to-speech feedback using pyttsx3
Configurable voice speed, volume and timeout settings

Mouse Control

Click, double-click, and right-click commands
Scroll up/down functionality
Mouse movement to specific coordinates

Keyboard Control

Press specific keys via voice
Type text directly using voice input

Window Management

Minimize/maximize the current window
Minimize all windows (show desktop)
Restore all windows
Individually control application windows

Application Control

Open applications by name
Close applications or the current window
Start and control Microsoft Office applications

Browser Control

Open websites with "browse to" command
Search using default search engine
Control tabs (open, close, navigate)
Navigation controls (back, forward, refresh)
Zoom controls and page interactions
Bookmark, save, and print pages

Microsoft Office Integration

Open specific Office applications (Word, Excel, PowerPoint, Outlook, OneNote)
Create new documents, spreadsheets, presentations, emails
Text formatting commands (bold, italic, underline)
Edit operations (copy, cut, paste, undo, redo)
Insert content (tables, images, slides)
Switch between viewing modes

GUI Details
Main Interface

Modern interface using customtkinter library
Light/Dark/System theme support
Split-pane design with sidebar and main content area

Sidebar

Logo/name display
Control buttons:

Start Listening
Stop Listening
Settings
Help
Exit


Appearance mode selector (Light/Dark/System)

Main Content Area

Status display showing current assistant state
Command log with timestamps
Command panel with tabbed organization:

Basic commands tab
Browser commands tab
Office commands tab



Settings Window

Tabbed interface for different settings categories:

General: Wake word and listening timeout
Browser: Default browser and search engine preferences
Voice: Voice selection, speed, and volume controls


Save button to apply configuration changes

Help System

Comprehensive help documentation
Tabbed interface organizing help by command category:

General commands and usage
Browser-specific commands
Microsoft Office commands


Command info tooltips when clicking on command buttons

The application has been designed with user-friendliness in mind, providing visual feedback about the assistant's status and commands while offering extensive voice control capabilities across the operating system, browsers, and productivity applications.

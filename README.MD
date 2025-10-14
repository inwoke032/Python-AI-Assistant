# Conversational AI Assistant (Self-Reprogrammable)

An intelligent, voice-controlled virtual assistant built in Python that can learn new skills on the fly by generating its own code. It operates in both English and Spanish and features a user-friendly GUI for configuration.

![Python Version](https://img.shields.io/badge/python-3.9+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

---

## 📸 Demo

![Assistant Screenshot](https://raw.githubusercontent.com/inwoke032/Python-AI-Assistant/blob/main/preview.gif)

---

## ✨ Key Features

-   **Voice Control**: Interact with your PC using natural voice commands in English or Spanish.
-   **Self-Reprogramming**: The assistant can learn new skills by generating, validating, and saving its own Python scripts.
-   **Multi-language Support**: Fully functional UI, voice recognition, and responses in both English and Spanish.
-   **Core Functions**:
    -   Open and close applications.
    -   Search on Google, YouTube, and Spotify.
    -   Control system volume and media playback (play/pause/next/previous).
    -   Get system status (CPU & RAM usage).
    -   Take screenshots.
-   **Persistent Memory**: Remembers user-specific facts (e.g., your name, hobbies) across sessions using a local SQLite database.
-   **User-Friendly Configuration**: A settings panel to easily change:
    -   Google Gemini API Key.
    -   UI and Voice Language.
    -   Text-to-Speech (TTS) voice.
    -   "Run on Startup" behavior for Windows.
-   **Special Modes**: Includes dedicated modes for real-time translation and note-taking.
-   **System Tray Integration**: Can be minimized to the system tray to run unobtrusively in the background.

---

## 🚀 Getting Started

Follow these instructions to get the assistant running on your local machine.

### Prerequisites

-   Python 3.9 or higher.
-   `pip` (Python package installer).
-   A Google Gemini API Key. You can get one for free from [Google AI Studio](https://aistudio.google.com/app/apikey).

### Installation Steps

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/inwoke032/Python-AI-Assistant.git
    cd Python-AI-Assistant
    ```

2.  **Create and activate a virtual environment (highly recommended):**
    ```bash
    # For Windows
    python -m venv venv
    .\venv\Scripts\activate

    # For macOS/Linux
    python3 -m venv venv
    source venv/bin/activate
    ```

3.  **Install the required dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Initial Configuration:**
    -   Run the application for the first time:
        ```bash
        python main.py
        ```
    -   Once the application launches, click the **⚙️ Settings** button.
    -   In the settings window, click **"Change API Key"** and paste your Google Gemini API key.
    -   Your key will be saved locally in `config.json` and is ready for use.

---

## 📖 Complete User Manual: How to Use the Assistant

This guide will walk you through every feature, from the basics to the most advanced capabilities.

### 1. Activation: How to "Wake Up" the Assistant

To give the assistant a command, you first need to get its attention. There are two simple methods:

* **Voice Activation (Wake Word)**:
    * Clearly say the phrase: **"Hey Assistant"** (for English) or **"Oye Asistente"** (for Spanish).
    * The application's interface will indicate that it is listening. You can then state your command. This is ideal for hands-free operation.

* **Manual Activation (Push-to-Talk)**:
    * Click the **"🎙️ Speak (PTT)"** button on the main window.
    * The button will change its state to show that the microphone is active. State your command. This is perfect for noisy environments or when you want full control over when the assistant listens.

### 2. Command and Skill Catalog

The assistant understands natural language, so you don't need to memorize exact phrases. Here is a comprehensive guide to its abilities with varied examples.

#### 2.1. Productivity and Access Commands

| Capability | Description | Example Commands |
| :--- | :--- | :--- |
| **Open Programs** | Launches any application installed on your PC. | "Open Google Chrome" <br> "Launch Spotify, please" <br> "Run calculator" |
| **Close Programs** | Terminates a running application's process. | "Close notepad" <br> "Terminate the Spotify process" |
| **Perform Calculations**| Solves simple mathematical operations. | "What is 125 times 8?" <br> "Calculate 1024 divided by 16" |

#### 2.2. Search and Information Commands

| Capability | Description | Example Commands |
| :--- | :--- | :--- |
| **Search Google** | Opens your browser to search for anything. | "Search for information about the history of computing" <br> "Google the recipe for lasagna" |
| **Search YouTube**| Finds and plays a video on YouTube. | "Play a video on YouTube about outer space" <br> "I want to watch a Python tutorial" |
| **Search Spotify**| Finds music in the Spotify application. | "Play music by Queen on Spotify" <br> "Search for the album 'Midnights'" |

#### 2.3. System and Media Control Commands

| Capability | Description | Example Commands |
| :--- | :--- | :--- |
| **System Status** | Reports the current CPU and RAM usage. | "What is the system status?" <br> "Tell me the PC's performance" |
| **Take Screenshot**| Saves a full-screen image to the program's folder. | "Take a screenshot" <br> "Capture the screen" |
| **Volume Control**| Modifies your system's master volume. | "Turn up the volume" <br> "Lower the volume" <br> "Mute" |
| **Media Control**| Controls playback in media players. | "Pause the music" <br> "Resume playing" <br> "Next song" <br> "Previous song" |

### 3. The Star Feature: Self-Learning

If you need the assistant to perform a task it doesn't know, you can teach it!

1.  **Trigger Learning**: Use the command **"Learn to..."** followed by the desired task.
2.  **Code Generation**: The assistant will use the Gemini AI to write a small Python script to perform the task.
3.  **Security Confirmation**: It will show you the generated code and ask for your permission to run it. **It is critical to read the code to ensure it is safe before you approve it.**
4.  **Execution and Saving**: If you approve, the script will run. If it succeeds, it will be saved as a new, permanent skill.

**Practical Example**:
> **You**: "Hey Assistant, learn to create a text file on the desktop named 'shopping list'."
>
> **Assistant**: "Understood. I have generated a script for this task. Do you want me to execute it?" *(Shows you the code)*.
>
> **You**: "Yes, go ahead."
>
> **Assistant**: "Done! I have learned the new skill and will remember it for the future."

### 4. Special Modes

The assistant can switch its behavior for specific tasks.

#### 4.1. Translator Mode

-   **To Activate**: "Activate translator mode to French" (or English, German, etc.).
-   **How it Works**: While active, anything you say will be translated into your chosen language. The assistant will speak the translation back to you.
-   **To Deactivate**: "Exit translator mode".

#### 4.2. Note-Taking Mode

-   **To Activate**: "Take a note" or "Write a note".
-   **How it Works**: Everything you say will be saved line-by-line into a `notes.txt` file in the program's folder, timestamped for your convenience.
-   **To Deactivate**: "End note".

---

## 🛠️ Tech Stack

-   **Language**: Python 3
-   **GUI**: Tkinter
-   **AI Model API**: Google Gemini
-   **Speech Recognition**: `speech_recognition` library
-   **Text-to-Speech (TTS)**: `pyttsx3`
-   **Local Database**: SQLite3
-   **System Automation**: `pyautogui`, `psutil`
-   **Windows Integration**: `winshell` (for "Run on Startup")

---

## 📄 License

This project is licensed under the MIT License. See the `LICENSE` file for details.

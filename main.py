# =============================================================================
# UPDATE NOTE (FINAL SOLUTION v15 - API KEY MANAGEMENT IN UI):
#
# The ability to view and change the Gemini API Key has been added
# directly from the user's graphical interface.
#
# 1. API KEY MANAGEMENT: A new button in the "Settings" window
#    allows the user to securely change their API Key.
#
# 2. DATA PERSISTENCE: The API Key is saved in the `config.json` file
#    so that it persists between sessions.
#
# 3. DYNAMIC UPDATE: The assistant starts using the new key
#    immediately after being saved, without needing to restart.
#
# 4. MORE ROBUST CODE: The global API Key variable was removed,
#    integrating it as an attribute of the VirtualAssistant class.
# =============================================================================

import speech_recognition as sr
import pyttsx3
import datetime
import webbrowser
import pywhatkit
import os
import subprocess
import time
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk, simpledialog
from threading import Thread
import queue
import wikipedia
import re
import psutil
import pyautogui
from deep_translator import GoogleTranslator
import json
import sys
import logging
import sqlite3
import requests
from urllib.parse import quote

# Logging handler to send records to the GUI queue
class QueueHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))

# --- MODEL AND DATABASE CONFIGURATION ---
GEMINI_MODEL = "gemini-2.5-flash-preview-05-20"
API_URL_BASE = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
DB_NAME = "assistant_memory.db"

try:
    from PIL import Image
    import pystray
except ImportError:
    Image, pystray = None, None
    print("WARNING: 'Pillow (PIL)' and 'pystray' libraries are missing.")

try:
    import winshell
except ImportError:
    winshell = None
    print("WARNING: 'winshell' library is missing. Startup functionality will be disabled.")

# Precompilation of regular expressions for Spanish
RE_LEARN_ES = re.compile(r'(?:aprende a|nueva habilidad para|enséñate a)\s+(.+)', flags=re.IGNORECASE)
RE_OPEN_ES = re.compile(r'(?:abre|lanza|ejecuta)\s+(.+)', flags=re.IGNORECASE)
RE_CLOSE_ES = re.compile(r'(?:cierra|termina)\s+(.+)', flags=re.IGNORECASE)
RE_SEARCH_GOOGLE_ES = re.compile(r'(?:busca|googlea|buscar|información de)\s+(.+)', flags=re.IGNORECASE)
RE_YOUTUBE_ES = re.compile(r'(?:youtube|pon un video|quiero ver)\s+(.+)', flags=re.IGNORECASE)
RE_SPOTIFY_ES = re.compile(r'(?:música|spotify|pon música|escuchar|reproduce)\s+(.+)', flags=re.IGNORECASE)
RE_CALCULATE_ES = re.compile(r'(?:calcula|cuánto es)\s+(.+)', flags=re.IGNORECASE)

# Precompilation of regular expressions for English
RE_LEARN_EN = re.compile(r'(?:learn to|new skill for|teach yourself to)\s+(.+)', flags=re.IGNORECASE)
RE_OPEN_EN = re.compile(r'(?:open|launch|run)\s+(.+)', flags=re.IGNORECASE)
RE_CLOSE_EN = re.compile(r'(?:close|terminate)\s+(.+)', flags=re.IGNORECASE)
RE_SEARCH_GOOGLE_EN = re.compile(r'(?:search|google|look for|information on)\s+(.+)', flags=re.IGNORECASE)
RE_YOUTUBE_EN = re.compile(r'(?:youtube|play a video|I want to watch)\s+(.+)', flags=re.IGNORECASE)
RE_SPOTIFY_EN = re.compile(r'(?:music|spotify|play music|listen to|play)\s+(.+)', flags=re.IGNORECASE)
RE_CALCULATE_EN = re.compile(r'(?:calculate|what is)\s+(.+)', flags=re.IGNORECASE)


class VirtualAssistant:
    CONFIG_FILE = "config.json"
    USER_CONFIG_FILE = "user_config.txt"
    HISTORY_FILE = "chat_history.json"
    DEFAULT_ASSISTANT_NAME_ES = "Asistente"
    DEFAULT_ASSISTANT_NAME_EN = "Assistant"
    SKILLS_DIR = "learned_skills"
    SKILLS_REGISTRY = "skills_registry.json"

    def __init__(self, app_instance):
        self.app = app_instance
        self.is_running = True
        self.config = self._load_configuration()
        self.language = self.config.get("language", "en") # Default to English
        
        self.assistant_name = self.config.get("assistant_name", self.DEFAULT_ASSISTANT_NAME_EN if self.language == "en" else self.DEFAULT_ASSISTANT_NAME_ES)
        self.tts_enabled = self.config.get("tts_enabled", True)
        self.api_key = self.config.get("api_key")
        self.user_name = self._load_user_name()
        self.user_id = self.user_name if self.user_name else "guest"
        self.translator_mode = False
        self.translation_language = None
        self.note_mode = False
        self.engine = None
        self.voices = []
        self.voice_index = self.config.get("voice_id", 0)
        self.tts_is_speaking = False
        self.wake_word_thread = None
        
        self._init_tts_thread()
        self._init_recognizer()
        self._init_db()
        self._update_language_settings()

        if not os.path.exists(self.SKILLS_DIR):
            os.makedirs(self.SKILLS_DIR)
        self.learned_skills = self._load_learned_skills()
        logging.info(f"Loaded {len(self.learned_skills)} learned skills.")

        if not self.api_key or self.api_key == "YOUR_API_KEY_HERE":
            logging.critical("SECURITY ALERT! The Gemini API Key is not configured. Change it in 'Settings'.")
    
    def _update_language_settings(self):
        """Sets language-specific variables."""
        self.WAKE_WORD = "hey assistant" if self.language == "en" else "oye asistente"
        wikipedia.set_lang(self.language)

        if self.language == 'en':
            self.command_registry = [
                {'regex': RE_LEARN_EN, 'handler': self.handle_learning_request},
                {'regex': RE_OPEN_EN, 'handler': self.open_application},
                {'regex': RE_CLOSE_EN, 'handler': self.close_application},
                {'regex': RE_SEARCH_GOOGLE_EN, 'handler': self.search_on_google},
                {'regex': RE_YOUTUBE_EN, 'handler': self.play_on_youtube},
                {'regex': RE_SPOTIFY_EN, 'handler': self.play_on_spotify},
                {'regex': RE_CALCULATE_EN, 'handler': self.calculate_arithmetic},
                {'keywords': ['system status', 'system information'], 'handler': self.system_status},
                {'keywords': ['screenshot', 'take a screenshot'], 'handler': self.take_screenshot},
                {'keywords': ['pause', 'play', 'next song', 'previous song', 'media control'], 'handler': self.control_media},
                {'keywords': ['volume up', 'volume down', 'mute'], 'handler': self.control_volume},
                {'keywords': ['translator to', 'translate to'], 'handler': self.start_translator_mode},
                {'keywords': ['take a note', 'write a note'], 'handler': self.start_note_mode},
                {'keywords': ['end note', 'finish note'], 'handler': self.end_note_mode},
            ]
        else: # Spanish
            self.command_registry = [
                {'regex': RE_LEARN_ES, 'handler': self.handle_learning_request},
                {'regex': RE_OPEN_ES, 'handler': self.open_application},
                {'regex': RE_CLOSE_ES, 'handler': self.close_application},
                {'regex': RE_SEARCH_GOOGLE_ES, 'handler': self.search_on_google},
                {'regex': RE_YOUTUBE_ES, 'handler': self.play_on_youtube},
                {'regex': RE_SPOTIFY_ES, 'handler': self.play_on_spotify},
                {'regex': RE_CALCULATE_ES, 'handler': self.calculate_arithmetic},
                {'keywords': ['estado del sistema', 'información del sistema'], 'handler': self.system_status},
                {'keywords': ['captura de pantalla', 'pantallazo'], 'handler': self.take_screenshot},
                {'keywords': ['pausa', 'reproduce', 'siguiente canción', 'anterior canción', 'control multimedia'], 'handler': self.control_media},
                {'keywords': ['sube el volumen', 'baja el volumen', 'silencio', 'mudo'], 'handler': self.control_volume},
                {'keywords': ['traductor al'], 'handler': self.start_translator_mode},
                {'keywords': ['tomar nota', 'escribe una nota'], 'handler': self.start_note_mode},
                {'keywords': ['terminar nota', 'finalizar nota'], 'handler': self.end_note_mode},
            ]


    def set_api_key(self, new_key):
        """Updates the API key and saves it to the configuration."""
        if new_key and isinstance(new_key, str):
            self.api_key = new_key
            self.save_configuration()
            logging.info("API Key updated successfully by the user.")
            return True
        return False

    def _load_configuration(self):
        default_api_key = os.getenv("GEMINI_API_KEY", "YOUR_API_KEY_HERE")
        config = {
            "assistant_name": self.DEFAULT_ASSISTANT_NAME_EN, 
            "tts_enabled": True, 
            "wake_word_enabled": True, 
            "voice_id": 0,
            "api_key": default_api_key,
            "language": "en" # Default language
        }
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config.update(json.load(f))
            except json.JSONDecodeError:
                logging.warning(f"JSON error in {self.CONFIG_FILE}.")
        try:
            config['voice_id'] = int(config['voice_id'])
        except (ValueError, TypeError):
            config['voice_id'] = 0
        return config

    def save_configuration(self):
        self.config.update({
            "assistant_name": self.assistant_name, 
            "tts_enabled": self.tts_enabled, 
            "wake_word_enabled": self.app.is_listening_continuously, 
            "voice_id": self.voice_index,
            "api_key": self.api_key,
            "language": self.language
        })
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            logging.error(f"Error saving configuration: {e}")

    def _call_gemini_api(self, user_query, include_grounding=False, structured_output=None):
        if not self.api_key or self.api_key == "YOUR_API_KEY_HERE":
            return "Error: Gemini API key is not configured. Please set it in the 'Settings' menu."
        
        system_prompt = (
            f"You are {self.assistant_name}, a friendly and helpful virtual assistant. "
            f"Purpose: Converse, answer questions, and execute commands.\n"
            f"--- CONTEXTUAL MEMORY ---\n{self._get_user_facts()}\n"
            f"--- INSTRUCTIONS ---\n1. Use known facts to personalize responses.\n"
            "2. Briefly acknowledge new personal data.\n3. Do not mention the database.\n4. Be concise."
        )
        payload = {"contents": [{"parts": [{"text": user_query}]}]}
        if "Act as an expert programmer" not in user_query and "Actúa como un programador experto" not in user_query:
            payload["systemInstruction"] = {"parts": [{"text": system_prompt}]}
        if include_grounding:
            payload['tools'] = [{"google_search": {}}]
        if structured_output:
            payload['generationConfig'] = {"responseMimeType": "application/json", "responseSchema": structured_output}
        
        try:
            for i in range(3):
                response = requests.post(f"{API_URL_BASE}?key={self.api_key}", headers={'Content-Type': 'application/json'}, json=payload, timeout=25)
                if response.status_code == 200:
                    result = response.json()
                    generated_text = result['candidates'][0]['content']['parts'][0]['text']
                    return json.loads(generated_text) if structured_output else generated_text
                elif response.status_code == 429:
                    time.sleep(2 ** i)
                    logging.warning(f"Rate limit reached. Retrying in {2 ** i}s...")
                else:
                    response.raise_for_status()
            return "The Gemini API did not respond. Check your connection."
        except Exception as e:
            logging.error(f"Error calling Gemini API: {e}")
            return "I can't connect to my brain. Check the connection and the API Key."

    def _extract_and_save_facts(self, user_query, assistant_response):
        fact_system_prompt = (
            "You are an information extractor. Analyze the conversation. If a personal fact about the user is mentioned (name, hobby, preference), "
            "extract it as a list of concise phrases. If there are no facts, return an empty JSON list: []."
        )
        conversation_context = f"User: '{user_query}'. Assistant: '{assistant_response}'."
        payload = {
            "contents": [{"parts": [{"text": conversation_context}]}],
            "systemInstruction": {"parts": [{"text": fact_system_prompt}]},
            "generationConfig": {"responseMimeType": "application/json", "responseSchema": {"type": "ARRAY", "items": {"type": "STRING"}}}
        }
        try:
            response = requests.post(f"{API_URL_BASE}?key={self.api_key}", headers={'Content-Type': 'application/json'}, json=payload, timeout=15)
            if response.status_code == 200:
                result = response.json()
                new_facts = json.loads(result['candidates'][0]['content']['parts'][0]['text'])
                if isinstance(new_facts, list) and new_facts:
                    conn = sqlite3.connect(DB_NAME)
                    cursor = conn.cursor()
                    for fact in new_facts:
                        cursor.execute("INSERT OR IGNORE INTO user_facts (user_id, fact) VALUES (?, ?)", (self.user_id, fact.strip()))
                        logging.info(f"Fact saved: {fact.strip()}")
                    conn.commit()
                    conn.close()
        except Exception as e:
            logging.error(f"Error in fact extraction: {e}")

    def _load_learned_skills(self):
        if os.path.exists(self.SKILLS_REGISTRY):
            try:
                with open(self.SKILLS_REGISTRY, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError) as e:
                logging.error(f"Error loading the skills registry: {e}")
                return {}
        return {}

    def _save_skill(self, command_key, code):
        file_name = f"skill_{re.sub(r'[^a-z0-9_]', '', command_key.lower().replace(' ', '_'))}.py"
        file_path = os.path.join(self.SKILLS_DIR, file_name)
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(code)
            self.learned_skills[command_key] = file_name
            with open(self.SKILLS_REGISTRY, 'w', encoding='utf-8') as f:
                json.dump(self.learned_skills, f, indent=4)
            logging.info(f"New skill '{command_key}' saved in '{file_name}'.")
            return f"Done! I've learned to '{command_key}' and will remember it for the future."
        except Exception as e:
            logging.error(f"Could not save the new skill: {e}")
            return "I was able to perform the action, but I had a problem saving it for the future."

    def handle_learning_request(self, command):
        match = (RE_LEARN_EN if self.language == 'en' else RE_LEARN_ES).search(command)
        if not match:
            return "I didn't understand what new skill you want me to learn. Please try again by saying 'learn to...' followed by the task."
        task = match.group(1).strip()
        return self.try_to_learn_skill(task)

    def try_to_learn_skill(self, task):
        self.app.add_text_to_chat(f"Understood. I will try to generate a script to learn how to '{task}'...", is_assistant=True)
        self.say_text(f"Understood. Let me see if I can learn to {task}.")
        
        programmer_prompt_en = (
            "Act as an expert Python programmer. Your task is to write a concise, self-contained Python script to perform a specific task on a Windows PC. "
            "Important rules:\n"
            "1. The script must be fully functional on its own.\n"
            "2. Use only standard Python libraries or the following pre-installed libraries: pyautogui, psutil, winshell, requests, webbrowser.\n"
            "3. DO NOT include code to install libraries (e.g., `pip install`).\n"
            "4. DO NOT define functions unless strictly necessary. Prefer a sequential script.\n"
            "5. The goal is to perform this task: '{}'.\n"
            "6. Return your response as a JSON object with a single key 'python_code' containing the code as a string. Do not add explanations outside the JSON.\n"
            "Example task: 'create a folder named tests on the desktop'.\n"
            'Example JSON response: {{"python_code": "import os\\nimport winshell\\ndesktop = winshell.desktop()\\nfolder_path = os.path.join(desktop, \\"tests\\")\\nos.makedirs(folder_path, exist_ok=True)"}}'
        ).format(task)
        
        programmer_prompt_es = (
            "Actúa como un programador experto de Python. Tu tarea es escribir un script de Python conciso y autocontenido para realizar una tarea específica en un PC con Windows. "
            "Reglas importantes:\n"
            "1. El script debe ser completamente funcional por sí mismo.\n"
            "2. Usa únicamente librerías estándar de Python o las siguientes librerías pre-instaladas: pyautogui, psutil, winshell, requests, webbrowser.\n"
            "3. NO incluyas código para instalar librerías (ej. `pip install`).\n"
            "4. NO definas funciones a menos que sea estrictamente necesario. Prefiere un script secuencial.\n"
            "5. El objetivo es realizar esta tarea: '{}'.\n"
            "6. Devuelve tu respuesta como un objeto JSON con una única clave 'python_code' que contenga el código como un string. No añadas explicaciones fuera del JSON.\n"
            "Ejemplo de tarea: 'crea una carpeta llamada pruebas en el escritorio'.\n"
            'Ejemplo de respuesta JSON: {{"python_code": "import os\\nimport winshell\\ndesktop = winshell.desktop()\\nfolder_path = os.path.join(desktop, \\"pruebas\\")\\nos.makedirs(folder_path, exist_ok=True)"}}'
        ).format(task)

        programmer_prompt = programmer_prompt_en if self.language == 'en' else programmer_prompt_es

        try:
            response_schema = {"type": "OBJECT", "properties": {"python_code": {"type": "STRING"}}}
            json_response = self._call_gemini_api(programmer_prompt, structured_output=response_schema)
            if not json_response or 'python_code' not in json_response:
                return "My attempt to generate code failed. I couldn't find a solution."
            
            generated_code = json_response['python_code']
            confirmation_message = (f"I have generated the following script to try '{task}'.\n\n--- CODE ---\n{generated_code}\n--------------\n\nWARNING: Running unknown code can be risky.\nDo you want me to execute it?")
            
            if not self.app.ask_user_confirmation(confirmation_message):
                return "Okay, I will not execute the code. Canceling the operation."
            
            self.app.add_text_to_chat("Confirmation received. Executing code...", is_assistant=False, tag='system')
            exec_globals = {'pyautogui': pyautogui, 'psutil': psutil, 'winshell': winshell, 'os': os, 'requests': requests, 'webbrowser': webbrowser, 're': re, 'time': time}
            
            try:
                exec(generated_code, exec_globals)
                return self._save_skill(task, generated_code)
            except Exception as e:
                logging.error(f"Error executing generated code for '{task}': {e}")
                return f"The code executed but failed with an error: {str(e)}. I have not learned the skill."
        except Exception as e:
            logging.error(f"Error in the skill learning process: {e}")
            return "An error occurred while trying to learn. Please check the logs."

    def _listen_for_wake_word_loop_sr(self):
        logging.info("Wake Word listening thread (SpeechRecognition) started.")
        ww_recognizer = sr.Recognizer()
        ww_recognizer.dynamic_energy_threshold = False
        ww_recognizer.energy_threshold = 1000
        ww_recognizer.pause_threshold = 0.5
        
        with sr.Microphone(sample_rate=16000) as source:
            logging.info(f"Listening in the background for the phrase: '{self.WAKE_WORD}'")
            while self.is_running and self.app.is_listening_continuously:
                if self.tts_is_speaking:
                    time.sleep(0.1)
                    continue
                try:
                    audio = ww_recognizer.listen(source, timeout=3, phrase_time_limit=4)
                    heard_text = ww_recognizer.recognize_google(audio, language=f"{self.language}-{self.language.upper()}").lower()
                    if self.WAKE_WORD in heard_text:
                        logging.info(f"Wake Word '{self.WAKE_WORD}' detected!")
                        self.app.root.after(0, self.app._on_wake_word_detected)
                        time.sleep(1)
                except sr.WaitTimeoutError: pass
                except sr.UnknownValueError: pass
                except sr.RequestError as e:
                    logging.error(f"Network error in wake word thread (Google Speech): {e}")
                    self.app.root.after(0, self.app._toggle_wake_word)
                    self.app.root.after(0, lambda: self.app.add_text_to_chat("Network error for Wake Word. Disabling continuous listening.", is_assistant=False, tag='system'))
                    break
                except Exception as e:
                    logging.error(f"Unexpected error in wake word thread: {e}")
                    time.sleep(1)
        logging.info("Wake Word listening thread stopped.")

    def _init_db(self):
        try:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("CREATE TABLE IF NOT EXISTS user_facts (user_id TEXT, fact TEXT, timestamp DATETIME DEFAULT CURRENT_TIMESTAMP, PRIMARY KEY (user_id, fact))")
            conn.commit()
            conn.close()
            logging.info(f"SQLite database '{DB_NAME}' initialized.")
        except sqlite3.Error as e:
            logging.error(f"Error initializing the database: {e}")
            self.app.add_text_to_chat("DB ERROR: Could not connect to local memory.", is_assistant=False, tag='system')

    def _get_user_facts(self):
        conn = None
        try:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("SELECT fact FROM user_facts WHERE user_id = ? ORDER BY timestamp DESC", (self.user_id,))
            facts = [row[0] for row in cursor.fetchall()]
            if facts:
                return f"Known facts about the user (total {len(facts)}): \n- {'\n- '.join(facts)}"
            return "No specific facts are known about the user."
        except sqlite3.Error as e:
            logging.error(f"Error getting facts: {e}")
            return "Error querying memory."
        finally:
            if conn: conn.close()

    def listen_for_command(self):
        if not self.microphone: return "error_not_understood"
        try:
            with self.microphone as source:
                self.app.add_text_to_chat("Listening for command...", is_assistant=False, tag='system')
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=10)
            self.app.add_text_to_chat("Processing command...", is_assistant=False, tag='system')
            return self.recognizer.recognize_google(audio, language=f"{self.language}-{self.language.upper()}").lower()
        except sr.WaitTimeoutError:
            self.app.add_text_to_chat("I didn't hear any command.", is_assistant=False, tag='system')
            return "timeout"
        except sr.UnknownValueError:
            self.app.add_text_to_chat("I couldn't understand the command.", is_assistant=False, tag='system')
            return "error_not_understood"
        except sr.RequestError as e:
            logging.error(f"Error with the transcription service: {e}")
            self.app.add_text_to_chat(f"Transcription service error: {e}", is_assistant=False, tag='system')
            return "error_service"
        except Exception as e:
            logging.error(f"!!! CRITICAL AUDIO ERROR (listening for command): {e} ({type(e).__name__})")
            return "error_unknown"

    def process_command(self, command):
        if not command or not isinstance(command, str): return None
        clean_command = re.sub(r'[¿?¡!]', '', command.lower()).strip()
        
        if command.startswith("error_") or command == "timeout":
            error_messages_en = {
                "error_not_understood": "I didn't understand. Can you repeat?", 
                "error_service": "Speech service failure.", 
                "error_unknown": "Unexpected error while listening.", 
                "timeout": ""
            }
            error_messages_es = {
                "error_no_entendido": "No te entendí. ¿Puedes repetirlo?", 
                "error_servicio": "Fallo en el servicio de voz.", 
                "error_desconocido": "Error inesperado al escuchar.", 
                "timeout": ""
            }
            error_message = (error_messages_en if self.language == 'en' else error_messages_es).get(command, "Error processing your voice.")
            return None if not error_message else self._call_gemini_api(f"Respond friendly: {error_message}")

        if self.translator_mode: return self.translate(command)
        if self.note_mode: return self.handle_note(command)

        for cmd_data in self.command_registry:
            if ('regex' in cmd_data and cmd_data['regex'].search(clean_command)) or \
               ('keywords' in cmd_data and any(kw in clean_command for kw in cmd_data['keywords'])):
                return cmd_data['handler'](command)
        
        for skill_key, script_name in self.learned_skills.items():
            if skill_key in clean_command:
                try:
                    full_path = os.path.join(self.SKILLS_DIR, script_name)
                    if not os.path.exists(full_path): continue
                    with open(full_path, 'r', encoding='utf-8') as f: code = f.read()
                    self.app.add_text_to_chat(f"Executing skill: '{skill_key}'...", is_assistant=False, tag='system')
                    exec(code, {'pyautogui': pyautogui, 'psutil': psutil, 'winshell': winshell, 'os': os, 'requests': requests, 'webbrowser': webbrowser, 're': re, 'time': time})
                    return f"Done, I executed the task '{skill_key}'."
                except Exception as e:
                    logging.error(f"Error executing skill '{skill_key}': {e}")
                    return f"I tried to use a learned skill, but it failed: {e}"
        
        logging.info(f"Treating as conversation: '{command}'")
        llm_response = self._call_gemini_api(command, include_grounding=True)
        Thread(target=self._extract_and_save_facts, args=(command, llm_response), daemon=True).start()
        return llm_response

    def start_note_mode(self, _):
        self.note_mode = True
        return "Note mode activated. Tell me what to write or 'end note'."

    def handle_note(self, command):
        if 'end note' in command or 'finish note' in command or 'terminar nota' in command or 'finalizar nota' in command:
            self.note_mode = False
            return "Note mode finished."
        try:
            with open("notes.txt", 'a', encoding='utf-8') as f:
                f.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] - {command.capitalize()}\n")
            return "Note saved. Continue or say 'end note'."
        except IOError: return "I couldn't save the note."

    def end_note_mode(self, _):
        self.note_mode = False
        return "Note mode finished."

    def play_on_spotify(self, command):
        match = (RE_SPOTIFY_EN if self.language == 'en' else RE_SPOTIFY_ES).search(command)
        query = match.group(1).strip() if match else ""
        if query:
            webbrowser.open(f"spotify:search:{quote(query)}")
            return f"Searching for '{query.capitalize()}' on Spotify."
        return "What do you want to listen to?"

    def _tts_worker(self):
        try: self.engine = pyttsx3.init('sapi5')
        except Exception as e: logging.critical(f"Could not initialize pyttsx3: {e}"); return
        
        self.engine.setProperty('rate', 160)
        self.voices = self.engine.getProperty('voices')
        if self.voices:
            if self.voice_index >= len(self.voices): self.voice_index = 0
            
            # Prioritize language-specific voice
            lang_keyword = 'english' if self.language == 'en' else 'spanish'
            for i, voice in enumerate(self.voices):
                if lang_keyword in voice.name.lower():
                    self.voice_index = i
                    break
            self.engine.setProperty('voice', self.voices[self.voice_index].id)
        
        self.engine.connect('finished-utterance', self.app._on_speech_finished)
        self.engine.startLoop(False)
        while self.is_running:
            try:
                task = self.tts_queue.get(block=False)
                if isinstance(task, str): 
                    self.tts_is_speaking = True
                    self.engine.say(task)
                elif isinstance(task, dict) and task.get('action') == 'change_voice':
                    new_index = task.get('index')
                    if self.voices and 0 <= new_index < len(self.voices): 
                        self.engine.setProperty('voice', self.voices[new_index].id)
                self.tts_queue.task_done()
            except queue.Empty: pass
            self.engine.iterate()
            time.sleep(0.1)
        self.engine.endLoop()

    def _init_tts_thread(self):
        self.tts_queue = queue.Queue()
        self.tts_thread = Thread(target=self._tts_worker, daemon=True)
        self.tts_thread.start()

    def _init_recognizer(self):
        self.recognizer = sr.Recognizer()
        self.recognizer.energy_threshold = 400
        self.recognizer.dynamic_energy_threshold = True
        self.recognizer.pause_threshold = 0.8
        try:
            self.microphone = sr.Microphone(sample_rate=16000, chunk_size=1024)
            with self.microphone as source: self.recognizer.adjust_for_ambient_noise(source, duration=1)
        except Exception as e:
            logging.critical(f"Could not initialize microphone: {e}")
            self.microphone = None

    def _save_user_name(self, name):
        try:
            with open(self.USER_CONFIG_FILE, 'w', encoding='utf-8') as f: f.write(name)
            self.user_name = name
            self.user_id = name
        except Exception as e: logging.error(f"Error saving user name: {e}")

    def _load_user_name(self):
        if os.path.exists(self.USER_CONFIG_FILE):
            try:
                with open(self.USER_CONFIG_FILE, 'r') as f: return f.read().strip()
            except Exception: return None
        return None

    def say_text(self, text):
        if self.tts_enabled and text: self.tts_queue.put(text)

    def open_application(self, command):
        match = (RE_OPEN_EN if self.language == 'en' else RE_OPEN_ES).search(command)
        if not match: return "Which application do you want me to open?"
        program = match.group(1).strip().lower()
        if 'spotify' in program:
            try: webbrowser.open("spotify://"); return "Opening **Spotify**."
            except Exception: program = "spotify.exe"
        try:
            subprocess.Popen([program])
            return f"Opening **{program.replace('.exe', '').capitalize()}**."
        except Exception:
            try:
                if not program.endswith('.exe'):
                    subprocess.Popen([program + '.exe'])
                    return f"Opening **{program.capitalize()}**."
                return f"I didn't find a program for **{program.capitalize()}**."
            except Exception: return f"I didn't find a program for **{program.capitalize()}**."

    def close_application(self, command):
        match = (RE_CLOSE_EN if self.language == 'en' else RE_CLOSE_ES).search(command)
        if not match: return "Which application do you want to close?"
        program = match.group(1).strip().lower()
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if program in proc.info['name'].lower():
                    proc.terminate()
                    proc.wait(timeout=3)
                    if proc.is_running(): proc.kill()
                    return f"**{program.capitalize()}** has been closed."
            except (psutil.NoSuchProcess, psutil.AccessDenied): continue
        return f"I couldn't find the process **{program}**."

    def control_media(self, command):
        if 'pause' in command or 'play' in command or 'reproduce' in command: pyautogui.press('playpause')
        elif 'next' in command or 'siguiente' in command: pyautogui.press('nexttrack')
        elif 'previous' in command or 'anterior' in command: pyautogui.press('prevtrack')
        return "Media command executed."

    def control_volume(self, command):
        if 'volume up' in command or 'sube el volumen' in command: pyautogui.press('volumeup')
        elif 'volume down' in command or 'baja el volumen' in command: pyautogui.press('volumedown')
        elif 'mute' in command or 'silencio' in command: pyautogui.press('volumemute')
        return "Volume control executed."

    def system_status(self, _):
        return f"CPU at **{psutil.cpu_percent(interval=1)}%**, RAM at **{psutil.virtual_memory().percent}%**."

    def take_screenshot(self, _):
        try:
            file_name = f"screenshot_{datetime.datetime.now():%Y%m%d_%H%M%S}.png"
            pyautogui.screenshot(file_name)
            return f"Screenshot saved as **{file_name}**."
        except Exception as e: return f"Error taking screenshot: {e}"

    def calculate_arithmetic(self, command):
        match = (RE_CALCULATE_EN if self.language == 'en' else RE_CALCULATE_ES).search(command)
        if not match: return "What do you want me to calculate?"
        expression = match.group(1).strip().replace('x', '*').replace('divided by', '/').replace('por', '*').replace('entre', '/')
        try: return f"The result is {eval(expression)}"
        except Exception: return "I couldn't perform the calculation."

    def play_on_youtube(self, command):
        match = (RE_YOUTUBE_EN if self.language == 'en' else RE_YOUTUBE_ES).search(command)
        query = match.group(1).strip() if match else ""
        if query:
            Thread(target=pywhatkit.playonyt, args=(query,), daemon=True).start()
            return f"Playing **{query}** on YouTube."
        return "What video would you like to watch?"

    def search_on_google(self, command):
        match = (RE_SEARCH_GOOGLE_EN if self.language == 'en' else RE_SEARCH_GOOGLE_ES).search(command)
        query = match.group(1).strip() if match else ""
        if query:
            webbrowser.open(f"https://www.google.com/search?q={quote(query)}")
            return f"Searching for **{query}** on Google."
        return "What do you want to search for?"

    LANGUAGE_MAP = {'english': 'en', 'inglés': 'en', 'spanish': 'es', 'español': 'es', 'french': 'fr', 'francés': 'fr', 'german': 'de', 'alemán': 'de'}
    def translate(self, text):
        if "exit translator mode" in text or "sal del modo traductor" in text:
            self.translator_mode = False; self.translation_language = None
            return "Exiting translator mode."
        try:
            return f"The translation is: {GoogleTranslator(source='auto', target=self.translation_language).translate(text=text)}"
        except Exception: return "Sorry, I couldn't translate that."

    def start_translator_mode(self, command):
        match = re.search(r'(?:to|al)\s+([a-zA-Záéíóú]+)', command, flags=re.IGNORECASE)
        language = match.group(1).lower() if match else None
        if language in self.LANGUAGE_MAP:
            self.translator_mode = True
            self.translation_language = self.LANGUAGE_MAP[language]
            return f"Translator mode activated to {language}. Tell me what to translate."
        return f"I don't recognize the language '{language}'. Try English, French, Spanish or German."

class App:
    def __init__(self, root):
        self.root = root
        self.console_queue = queue.Queue()
        self._configure_logging()
        
        self.assistant = VirtualAssistant(self)
        self.language = self.assistant.language

        self._setup_window()
        self.is_listening_continuously = self.assistant.config.get("wake_word_enabled", True)
        self.is_running = True
        self.tray_icon = None
        self.chat_history = []
        
        self._setup_ui()
        self._setup_tray_icon()
        self._poll_console_queue()

        self._load_history()
        if self.is_listening_continuously:
            self._start_continuous_listening()
        if not self.chat_history or self.chat_history[-1]['tag'] == 'system':
            greeting = "Hello, briefly introduce yourself and explain that you are learning over time." if self.language == 'en' else "Hola, preséntate brevemente y explica que estás aprendiendo con el tiempo."
            self._handle_assistant_response_root(self.assistant._call_gemini_api(greeting))

    def _configure_logging(self):
        log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler = logging.FileHandler("assistant.log", encoding='utf-8')
        file_handler.setFormatter(log_formatter)
        queue_handler = QueueHandler(self.console_queue)
        queue_handler.setFormatter(log_formatter)
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        if logger.hasHandlers():
            logger.handlers.clear()
        logger.addHandler(file_handler)
        logger.addHandler(queue_handler)

    def _poll_console_queue(self):
        while not self.console_queue.empty():
            try:
                record = self.console_queue.get(block=False)
                self._update_console_text(record)
            except queue.Empty:
                pass
        self.root.after(100, self._poll_console_queue)

    def _update_console_text(self, text):
        self.console_area.config(state='normal')
        tag = "info"
        if "ERROR" in text or "CRITICAL" in text:
            tag = "error"
        elif "WARNING" in text:
            tag = "warning"
        self.console_area.insert('end', text + '\n', tag)
        self.console_area.config(state='disabled')
        self.console_area.see('end')

    def ask_user_confirmation(self, message):
        title = "Code Execution Confirmation"
        return messagebox.askyesno(title, message, parent=self.root)

    def _setup_window(self):
        self.root.title("Conversational Assistant (Self-Reprogrammable)")
        self.root.geometry("700x750")
        self.root.minsize(500, 600)
        self.root.configure(bg="#121212")
        self.root.protocol("WM_DELETE_WINDOW", self._handle_close)

    def _load_history(self):
        if os.path.exists(VirtualAssistant.HISTORY_FILE):
            try:
                with open(VirtualAssistant.HISTORY_FILE, 'r', encoding='utf-8') as f:
                    self.chat_history = json.load(f)
                for entry in self.chat_history:
                    self._add_history_text(entry)
                self.add_text_to_chat("Chat history loaded.", is_assistant=False, tag='system')
            except Exception as e:
                logging.error(f"Could not load history: {e}")

    def _add_history_text(self, entry):
        self.chat_area.config(state='normal')
        tag = entry.get('tag', 'assistant' if entry['is_assistant'] else 'user')
        you_text = "You" if self.language == "en" else "Tú"
        prefix = f"[{entry.get('timestamp', '')}] {self.assistant.assistant_name if entry['is_assistant'] else you_text}: "
        self.chat_area.insert('end', prefix, 'system')
        self.chat_area.insert('end', entry['text'] + '\n\n', tag)
        self.chat_area.config(state='disabled')
        self.chat_area.see('end')

    def _save_history(self):
        try:
            with open(VirtualAssistant.HISTORY_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.chat_history[-100:], f, indent=4)
        except Exception as e:
            logging.error(f"Could not save history: {e}")

    def _setup_tray_icon(self):
        if pystray and Image:
            try:
                image = Image.open("icon.png") # Assume you have an icon.png
            except FileNotFoundError:
                image = Image.new('RGB', (64, 64), color='#1e1e1e')
            
            show_text = "Show" if self.language == 'en' else "Mostrar"
            exit_text = "Exit" if self.language == 'en' else "Salir"
            menu = (pystray.MenuItem(show_text, self._show_window), pystray.MenuItem(exit_text, self._handle_close))
            self.tray_icon = pystray.Icon("Assistant", image, "Virtual Assistant", menu)

    def _hide_window(self):
        if self.tray_icon:
            self.root.withdraw()
            Thread(target=self.tray_icon.run, daemon=True).start()
        else:
            self.root.iconify()
        self.add_text_to_chat(f"{self.assistant.assistant_name} in the background.", is_assistant=False, tag='system')

    def _show_window(self, icon=None, item=None):
        if self.tray_icon and icon:
            icon.stop()
        self.root.after(0, self.root.deiconify)
        self.root.after(100, self.root.lift)

    def _open_settings_window(self):
        settings_win = tk.Toplevel(self.root)
        settings_win.title("Advanced Settings")
        settings_win.geometry("450x450")
        BG_COLOR, FG_COLOR = "#212121", "#e0e0e0"
        ACCENT_COLOR, BUTTON_COLOR = "#61dafb", "#424242"
        settings_win.configure(bg=BG_COLOR)

        tk.Label(settings_win, text="Configuration Options", font=("Segoe UI", 14, "bold"), bg=BG_COLOR, fg=ACCENT_COLOR).pack(pady=10)

        # Language selection
        lang_frame = tk.LabelFrame(settings_win, text="Language", font=("Segoe UI", 11, "bold"), bg=BG_COLOR, fg=FG_COLOR, padx=10, pady=10)
        lang_frame.pack(pady=10, padx=10, fill='x')
        
        self.lang_var = tk.StringVar(value=self.assistant.language)
        
        def on_lang_change():
            self.assistant.language = self.lang_var.get()
            self.assistant.save_configuration()
            messagebox.showinfo("Language Change", "Language updated. Please restart the application for all changes to take effect.", parent=settings_win)

        tk.Radiobutton(lang_frame, text="English", variable=self.lang_var, value="en", bg=BG_COLOR, fg=FG_COLOR, selectcolor=BUTTON_COLOR, command=on_lang_change).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(lang_frame, text="Español", variable=self.lang_var, value="es", bg=BG_COLOR, fg=FG_COLOR, selectcolor=BUTTON_COLOR, command=on_lang_change).pack(side=tk.LEFT, padx=10)

        # Voice and Audio
        voice_frame = tk.LabelFrame(settings_win, text="Voice & Audio (TTS)", font=("Segoe UI", 11, "bold"), bg=BG_COLOR, fg=FG_COLOR, padx=10, pady=10)
        voice_frame.pack(pady=10, padx=10, fill='x')
        voice_label = tk.Label(voice_frame, text=f"Current Voice ID: {self.assistant.voice_index + 1}", bg=BG_COLOR, fg=FG_COLOR)
        voice_label.pack(side=tk.LEFT, padx=5)
        
        def update_voice():
            if self.assistant.voices:
                self.assistant.voice_index = (self.assistant.voice_index + 1) % len(self.assistant.voices)
                self.assistant.tts_queue.put({'action': 'change_voice', 'index': self.assistant.voice_index})
                voice_label.config(text=f"Current Voice ID: {self.assistant.voice_index + 1}")
        
        tk.Button(voice_frame, text="Change Voice", command=update_voice, bg=BUTTON_COLOR, fg=ACCENT_COLOR, relief="flat").pack(side=tk.RIGHT, padx=5)

        # Credentials
        api_frame = tk.LabelFrame(settings_win, text="Credentials", font=("Segoe UI", 11, "bold"), bg=BG_COLOR, fg=FG_COLOR, padx=10, pady=10)
        api_frame.pack(pady=10, padx=10, fill='x')
        tk.Label(api_frame, text="Gemini API Key:", bg=BG_COLOR, fg=FG_COLOR).pack(side=tk.LEFT, padx=5)
        tk.Button(api_frame, text="Change API Key", command=self._change_api_key_dialog, bg=BUTTON_COLOR, fg=ACCENT_COLOR, relief="flat").pack(side=tk.RIGHT, padx=5)

        # Startup option
        startup_frame = tk.LabelFrame(settings_win, text="System Integration", font=("Segoe UI", 11, "bold"), bg=BG_COLOR, fg=FG_COLOR, padx=10, pady=10)
        startup_frame.pack(pady=10, padx=10, fill='x')
        self.startup_var = tk.BooleanVar(value=self._is_startup_enabled())
        startup_check = tk.Checkbutton(startup_frame, text="Run on Startup", variable=self.startup_var, bg=BG_COLOR, fg=FG_COLOR, selectcolor=BUTTON_COLOR, command=self._toggle_startup)
        startup_check.pack(side=tk.LEFT)
        if not winshell:
            startup_check.config(state=tk.DISABLED)

        tk.Label(settings_win, text="*Conversational intelligence is fed by local memory.*", font=("Segoe UI", 10), bg=BG_COLOR, fg="#FFC107").pack(padx=20, pady=20)
        settings_win.transient(self.root)
        settings_win.grab_set()
        self.root.wait_window(settings_win)

    def _get_startup_shortcut_path(self):
        """Gets the path for the startup shortcut."""
        if not winshell:
            return None
        return os.path.join(winshell.startup(), "ConversationalAssistant.lnk")

    def _is_startup_enabled(self):
        """Checks if the startup shortcut exists."""
        shortcut_path = self._get_startup_shortcut_path()
        return shortcut_path and os.path.exists(shortcut_path)

    def _toggle_startup(self):
        """Adds or removes the application from Windows startup."""
        if not winshell:
            logging.warning("winshell is not available, cannot manage startup.")
            return

        shortcut_path = self._get_startup_shortcut_path()
        target_path = sys.executable
        working_dir = os.path.dirname(target_path)

        if self.startup_var.get(): # If checkbox is checked
            try:
                with winshell.shortcut(shortcut_path) as shortcut:
                    shortcut.path = target_path
                    shortcut.working_directory = working_dir
                    shortcut.description = "Conversational Assistant"
                logging.info("Startup shortcut created.")
                self.add_text_to_chat("Enabled to run on startup.", is_assistant=False, tag='system')
            except Exception as e:
                logging.error(f"Failed to create startup shortcut: {e}")
                messagebox.showerror("Startup Error", f"Could not create startup shortcut: {e}", parent=self.root)
        else: # If checkbox is unchecked
            try:
                if os.path.exists(shortcut_path):
                    os.remove(shortcut_path)
                    logging.info("Startup shortcut removed.")
                    self.add_text_to_chat("Disabled from running on startup.", is_assistant=False, tag='system')
            except Exception as e:
                logging.error(f"Failed to remove startup shortcut: {e}")
                messagebox.showerror("Startup Error", f"Could not remove startup shortcut: {e}", parent=self.root)

    def _change_api_key_dialog(self):
        current_key = self.assistant.api_key
        masked_key = f"{current_key[:4]}...{current_key[-4:]}" if current_key and len(current_key) > 8 else "Not set"

        new_key = simpledialog.askstring(
            "Change API Key",
            f"Enter your new Google Gemini API Key.\n\nCurrent key: {masked_key}",
            parent=self.root
        )

        if new_key and new_key.strip():
            if self.assistant.set_api_key(new_key.strip()):
                self.add_text_to_chat("API Key updated successfully.", is_assistant=False, tag='system')
            else:
                messagebox.showerror("Error", "Could not update the API Key.", parent=self.root)
        else:
            logging.info("API Key change canceled by the user.")

    def _setup_ui(self):
        BG_COLOR, FG_COLOR = "#1e1e1e", "#e0e0e0"
        CHAT_BG, BUTTON_COLOR = "#212121", "#323232"
        ACCENT_BLUE, ACCENT_RED, ACCENT_GREEN = "#61dafb", "#F44336", "#4CAF50"
        FONT_NAME, FONT_SIZE = "Segoe UI", 11

        style = ttk.Style()
        style.configure("TNotebook", background=BG_COLOR, borderwidth=0)
        style.configure("TNotebook.Tab", background=BUTTON_COLOR, foreground=FG_COLOR, padding=[10, 5], font=(FONT_NAME, 10, "bold"))
        style.map("TNotebook.Tab", background=[("selected", CHAT_BG)])

        notebook = ttk.Notebook(self.root, style="TNotebook")
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Translations for UI elements
        assistant_tab_text = "Assistant" if self.language == "en" else "Asistente"
        console_tab_text = "Console" if self.language == "en" else "Consola"
        send_button_text = "Send" if self.language == "en" else "Enviar"
        ptt_button_text = "🎙️ Speak (PTT)" if self.language == "en" else "🎙️ Hablar (PTT)"
        voice_on_text = "🔊 Voice ON" if self.language == "en" else "🔊 Voz ON"
        voice_off_text = "🔇 Voice OFF" if self.language == "en" else "🔇 Voz OFF"
        settings_button_text = "⚙️ Settings" if self.language == "en" else "⚙️ Config"
        hide_button_text = "➖ Hide" if self.language == "en" else "➖ Ocultar"
        wake_word_on_text = "Wake Word: ON" if self.language == "en" else "Wake Word: ON" # Same for both for simplicity
        wake_word_off_text = "Wake Word: OFF" if self.language == "en" else "Wake Word: OFF"

        chat_tab_frame = ttk.Frame(notebook, style="TFrame")
        notebook.add(chat_tab_frame, text=assistant_tab_text)
        main_frame = tk.Frame(chat_tab_frame, bg=BG_COLOR, padx=15, pady=15)
        main_frame.pack(fill="both", expand=True)

        self.chat_area = scrolledtext.ScrolledText(main_frame, state='disabled', wrap='word', font=(FONT_NAME, FONT_SIZE), bg=CHAT_BG, fg=FG_COLOR, bd=0, relief="flat", padx=15, pady=15)
        self.chat_area.pack(fill="both", expand=True, pady=(0, 10))
        self.chat_area.tag_config('assistant', foreground=ACCENT_BLUE, font=(FONT_NAME, FONT_SIZE, "bold"))
        self.chat_area.tag_config('user', foreground='#b0b0b0')
        self.chat_area.tag_config('system', foreground='#FFC107', font=(FONT_NAME, 10, "italic"))
        self.chat_area.tag_config('user_invoked', foreground=FG_COLOR)
        
        input_frame = tk.Frame(main_frame, bg=BG_COLOR)
        input_frame.pack(fill="x", pady=(0, 10))
        self.text_entry = tk.Entry(input_frame, font=(FONT_NAME, 12), bg=CHAT_BG, fg=FG_COLOR, bd=0, relief="flat", insertbackground=FG_COLOR)
        self.text_entry.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5), ipady=8)
        self.text_entry.bind("<Return>", self._process_text_entry)
        tk.Button(input_frame, text=send_button_text, command=lambda: self._process_text_entry(None), font=(FONT_NAME, 10, "bold"), bg=ACCENT_BLUE, fg="white", relief="flat", padx=15, pady=5).pack(side=tk.RIGHT)
        
        ptt_frame = tk.Frame(main_frame, bg=BG_COLOR)
        ptt_frame.pack(fill="x", pady=(5, 5))
        self.listen_button = tk.Button(ptt_frame, text=ptt_button_text, command=self._start_listening_thread, font=(FONT_NAME, 12, "bold"), bg=BUTTON_COLOR, fg="white", relief="flat", padx=15, pady=10)
        self.listen_button.pack(side=tk.LEFT, expand=True, fill="x")
        
        toggle_frame = tk.Frame(main_frame, bg=BG_COLOR)
        toggle_frame.pack(pady=(5, 0), fill="x")
        
        self.tts_button = tk.Button(toggle_frame, text=voice_on_text, command=self._toggle_tts, font=(FONT_NAME, 10, "bold"), bg=ACCENT_GREEN, fg="white", relief="flat", padx=10, pady=5)
        self.tts_button.pack(side=tk.LEFT, padx=(0, 5), expand=True)
        if not self.assistant.tts_enabled: self.tts_button.config(text=voice_off_text, bg=ACCENT_RED)

        self.config_button = tk.Button(toggle_frame, text=settings_button_text, command=self._open_settings_window, font=(FONT_NAME, 10, "bold"), bg=BUTTON_COLOR, fg=FG_COLOR, relief="flat", padx=10, pady=5)
        self.config_button.pack(side=tk.LEFT, padx=5, expand=True)
        
        self.hide_button = tk.Button(toggle_frame, text=hide_button_text, command=self._hide_window, font=(FONT_NAME, 10, "bold"), bg=BUTTON_COLOR, fg=FG_COLOR, relief="flat", padx=10, pady=5)
        self.hide_button.pack(side=tk.LEFT, padx=5, expand=True)

        wake_color = ACCENT_BLUE if self.is_listening_continuously else BUTTON_COLOR
        wake_status = wake_word_on_text if self.is_listening_continuously else wake_word_off_text
        self.wake_word_button = tk.Button(toggle_frame, text=wake_status, command=self._toggle_wake_word, font=(FONT_NAME, 10, "bold"), bg=wake_color, fg="white", relief="flat", padx=10, pady=5)
        self.wake_word_button.pack(side=tk.RIGHT, padx=(5, 0), expand=True)
        
        console_tab_frame = ttk.Frame(notebook, style="TFrame")
        notebook.add(console_tab_frame, text=console_tab_text)
        self.console_area = scrolledtext.ScrolledText(console_tab_frame, wrap='char', font=("Consolas", 9), bg="#101010", fg="#c0c0c0", bd=0, relief="flat")
        self.console_area.pack(fill="both", expand=True, padx=5, pady=5)
        self.console_area.tag_config('info', foreground="#c0c0c0")
        self.console_area.tag_config('warning', foreground="#FFC107")
        self.console_area.tag_config('error', foreground="#F44336")
        self.console_area.config(state='disabled')

    def _process_text_entry(self, event=None):
        command = self.text_entry.get().strip()
        if not command: return
        self.text_entry.delete(0, tk.END)
        self.add_text_to_chat(command, is_assistant=False)
        Thread(target=self._execute_logic_in_thread, args=(command,), daemon=True).start()

    def _on_speech_finished(self, name, completed):
        self.root.after(0, self._on_speech_finished_thread_safe, completed)

    def _on_speech_finished_thread_safe(self, completed):
        if completed:
            self.assistant.tts_is_speaking = False
            logging.info("TTS has finished speaking.")
            self._reactivate_button()

    def _on_wake_word_detected(self):
        self.root.lift()
        self.root.focus_force()
        self.listen_button.config(state=tk.DISABLED, text="Detected...")
        self.add_text_to_chat(f"{self.assistant.assistant_name} is listening!", is_assistant=False, tag='system')
        Thread(target=lambda: self._execute_voice_logic_in_thread(was_by_wake_word=True), daemon=True).start()

    def _start_continuous_listening(self):
        self.is_listening_continuously = True
        self.assistant.config["wake_word_enabled"] = True
        self.assistant.save_configuration()
        self.assistant.wake_word_thread = Thread(target=self.assistant._listen_for_wake_word_loop_sr, daemon=True)
        self.assistant.wake_word_thread.start()

    def _stop_continuous_listening(self):
        self.is_listening_continuously = False
        self.assistant.config["wake_word_enabled"] = False
        self.assistant.save_configuration()
        self.add_text_to_chat("Continuous listening disabled.", is_assistant=False, tag='system')

    def _toggle_wake_word(self):
        wake_word_on_text = "Wake Word: ON" if self.language == "en" else "Wake Word: ON"
        wake_word_off_text = "Wake Word: OFF" if self.language == "en" else "Wake Word: OFF"
        if self.is_listening_continuously:
            self._stop_continuous_listening()
            self.wake_word_button.config(text=wake_word_off_text, bg="#323232")
        else:
            self._start_continuous_listening()
            self.wake_word_button.config(text=wake_word_on_text, bg="#61dafb")

    def _toggle_tts(self):
        self.assistant.tts_enabled = not self.assistant.tts_enabled
        self.assistant.save_configuration()
        
        voice_on_text = "🔊 Voice ON" if self.language == "en" else "🔊 Voz ON"
        voice_off_text = "🔇 Voice OFF" if self.language == "en" else "🔇 Voz OFF"
        
        if self.assistant.tts_enabled:
            self.tts_button.config(text=voice_on_text, bg="#4CAF50")
            status_msg = "Voice (TTS) enabled."
        else:
            self.tts_button.config(text=voice_off_text, bg="#F44336")
            status_msg = "Voice (TTS) disabled."
        
        self.add_text_to_chat(status_msg, is_assistant=False, tag='system')

    def add_text_to_chat(self, text, is_assistant=True, tag=None):
        self.root.after(0, self._add_text, text, is_assistant, tag)

    def _add_text(self, text, is_assistant, tag):
        self.chat_area.config(state='normal')
        if tag is None: tag = 'assistant' if is_assistant else 'user'
        timestamp = datetime.datetime.now().strftime('%H:%M:%S')
        
        you_text = "You" if self.language == "en" else "Tú"
        prefix = f"[{timestamp}] {self.assistant.assistant_name if is_assistant else you_text}: "
        
        self.chat_area.insert('end', prefix, 'system')
        self.chat_area.insert('end', str(text) + '\n\n', tag)
        if tag != 'system':
            self.chat_history.append({'text': str(text), 'is_assistant': is_assistant, 'tag': tag, 'timestamp': timestamp})
        self.chat_area.config(state='disabled')
        self.chat_area.see('end')

    def _start_listening_thread(self):
        self.listen_button.config(state=tk.DISABLED, text="Listening...")
        Thread(target=lambda: self._execute_voice_logic_in_thread(was_by_wake_word=False), daemon=True).start()

    def _execute_voice_logic_in_thread(self, was_by_wake_word=True):
        command = self.assistant.listen_for_command()
        if command and "error_" not in command and "timeout" not in command:
            tag = 'user_invoked' if was_by_wake_word else 'user'
            self.root.after(0, lambda: self.add_text_to_chat(command, is_assistant=False, tag=tag))
        response = self.assistant.process_command(command)
        self._handle_assistant_response_root(response)

    def _execute_logic_in_thread(self, command):
        response = self.assistant.process_command(command)
        self._handle_assistant_response_root(response)

    def _handle_assistant_response_root(self, text):
        self.root.after(0, self.__handle_assistant_response_sync, text)

    def __handle_assistant_response_sync(self, text):
        if text:
            self.add_text_to_chat(text)
            self.assistant.say_text(text)
        self._reactivate_button()

    def _reactivate_button(self):
        if not self.assistant.tts_is_speaking:
            ptt_button_text = "🎙️ Speak (PTT)" if self.language == "en" else "🎙️ Hablar (PTT)"
            self.listen_button.config(state=tk.NORMAL, text=ptt_button_text)

    def _handle_close(self):
        logging.info("Initiating closing sequence...")
        self.is_running = False
        self.assistant.is_running = False
        if self.assistant.wake_word_thread and self.assistant.wake_word_thread.is_alive():
            self.assistant.wake_word_thread.join(timeout=1)
        self.assistant.save_configuration()
        self._save_history()
        if self.tray_icon: self.tray_icon.stop()
        self.root.destroy()
        logging.info("Application closed successfully.")
        os._exit(0)

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = App(root)
        root.mainloop()
    except KeyboardInterrupt:
        if 'app' in locals(): app._handle_close()
    except Exception as e:
        logging.critical(f"Fatal exception in main thread: {e}", exc_info=True)
        print(f"FATAL ERROR: {e}")
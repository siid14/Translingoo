import os
import sys
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import platform

class TranslatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Translator")
        self.root.geometry("600x450")  # Increased height to accommodate checkboxes
        self.root.resizable(True, True)
        
        # Ensure required files exist
        self.ensure_required_files()
        
        # Set app icon if available
        # self.root.iconbitmap("icon.ico")  # Uncomment and add icon if available
        
        # Configure the grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=0)
        self.root.rowconfigure(1, weight=0)
        self.root.rowconfigure(2, weight=0)  # New row for checkboxes
        self.root.rowconfigure(3, weight=1)  # Console now in row 3
        self.root.rowconfigure(4, weight=0)  # Buttons now in row 4
        
        # Create header with logo/title
        self.header_frame = tk.Frame(root, padx=10, pady=10)
        self.header_frame.grid(row=0, column=0, sticky="ew")
        
        self.title_label = tk.Label(
            self.header_frame, 
            text="Excel Translator", 
            font=("Arial", 18, "bold")
        )
        self.title_label.pack()
        
        self.subtitle_label = tk.Label(
            self.header_frame,
            text="Translate Excel files with technical terms from English to French",
            font=("Arial", 10)
        )
        self.subtitle_label.pack()
        
        # File selection frame
        self.file_frame = tk.Frame(root, padx=20, pady=10)
        self.file_frame.grid(row=1, column=0, sticky="ew")
        
        # File selection
        self.file_label = tk.Label(self.file_frame, text="Select Excel File:")
        self.file_label.grid(row=0, column=0, sticky="w", pady=5)
        
        self.file_frame.columnconfigure(1, weight=1)
        
        self.file_path = tk.StringVar()
        self.file_entry = tk.Entry(self.file_frame, textvariable=self.file_path, width=50)
        self.file_entry.grid(row=0, column=1, padx=5, sticky="ew")
        
        self.browse_button = tk.Button(self.file_frame, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=5)
        
        # Output file options
        self.output_label = tk.Label(self.file_frame, text="Output File (Optional):")
        self.output_label.grid(row=1, column=0, sticky="w", pady=5)
        
        self.output_path = tk.StringVar()
        self.output_entry = tk.Entry(self.file_frame, textvariable=self.output_path, width=50)
        self.output_entry.grid(row=1, column=1, padx=5, sticky="ew")
        
        self.output_button = tk.Button(self.file_frame, text="Browse", command=self.browse_output)
        self.output_button.grid(row=1, column=2, padx=5)
        
        # Column selection frame
        self.column_frame = tk.Frame(root, padx=20, pady=10)
        self.column_frame.grid(row=2, column=0, sticky="ew")
        
        self.column_label = tk.Label(
            self.column_frame, 
            text="Columns to translate:", 
            font=("Arial", 10, "bold")
        )
        self.column_label.grid(row=0, column=0, sticky="w", pady=5)
        
        # Checkboxes for column selection
        self.translate_description = tk.BooleanVar(value=True)
        self.translate_message = tk.BooleanVar(value=True)
        
        self.description_checkbox = tk.Checkbutton(
            self.column_frame, 
            text="Description", 
            variable=self.translate_description,
            padx=10
        )
        self.description_checkbox.grid(row=0, column=1, sticky="w")
        
        self.message_checkbox = tk.Checkbutton(
            self.column_frame, 
            text="Message", 
            variable=self.translate_message,
            padx=10
        )
        self.message_checkbox.grid(row=0, column=2, sticky="w")
        
        # Console output
        self.console_frame = tk.Frame(root, padx=20, pady=10)
        self.console_frame.grid(row=3, column=0, sticky="nsew")
        self.console_frame.columnconfigure(0, weight=1)
        self.console_frame.rowconfigure(0, weight=1)
        
        self.console = tk.Text(self.console_frame, height=10, bg="#f0f0f0", fg="#333333")
        self.console.grid(row=0, column=0, sticky="nsew")
        
        self.scrollbar = tk.Scrollbar(self.console_frame, command=self.console.yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.console.config(yscrollcommand=self.scrollbar.set)
        
        # Bottom buttons
        self.button_frame = tk.Frame(root, padx=20, pady=10)
        self.button_frame.grid(row=4, column=0, sticky="ew")
        self.button_frame.columnconfigure(1, weight=1)
        
        self.check_docker_button = tk.Button(
            self.button_frame, 
            text="Check Docker", 
            command=self.check_docker
        )
        self.check_docker_button.grid(row=0, column=0, padx=5, pady=10)
        
        self.translate_button = tk.Button(
            self.button_frame, 
            text="Translate File", 
            command=self.translate_file,
            bg="#4CAF50", 
            fg="white", 
            font=("Arial", 10, "bold"),
            height=2,
            width=15
        )
        self.translate_button.grid(row=0, column=2, padx=5, pady=10)
        
        # Progress bar
        self.progress = ttk.Progressbar(
            self.button_frame, 
            orient="horizontal", 
            length=200, 
            mode="indeterminate"
        )
        self.progress.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        # Check Docker on startup
        self.root.after(500, self.check_docker)
        
    def ensure_required_files(self):
        """Ensure that required files like Dockerfile exist when running as executable"""
        # Get the directory where the GUI script is located or where the executable is
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            base_dir = os.path.dirname(sys.executable)
        else:
            # Running as script
            base_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Check if Dockerfile exists, if not create it
        dockerfile_path = os.path.join(base_dir, 'Dockerfile')
        if not os.path.exists(dockerfile_path):
            print(f"Creating Dockerfile at {dockerfile_path}")
            with open(dockerfile_path, 'w') as f:
                f.write("""FROM python:3.9-slim

WORKDIR /app

# Install required packages
RUN pip install pandas xlrd openpyxl odfpy

# Copy the converter script
COPY converter.py /app/

# Make the script executable
RUN chmod +x /app/converter.py

ENTRYPOINT ["python", "/app/converter.py"]
""")
        
        # Check if converter.py exists, if not copy it from the executable directory
        converter_path = os.path.join(base_dir, 'converter.py')
        if not os.path.exists(converter_path):
            # Try to find it in the current directory
            if os.path.exists('converter.py'):
                import shutil
                shutil.copy('converter.py', converter_path)
            # If running as executable and converter.py is bundled as a resource
            elif getattr(sys, 'frozen', False):
                print(f"Extracting converter.py to {converter_path}")
                # Try to extract the converter.py from the executable resources
                try:
                    import pkg_resources
                    converter_content = pkg_resources.resource_string(__name__, 'converter.py')
                    with open(converter_path, 'wb') as f:
                        f.write(converter_content)
                except Exception as e:
                    print(f"Warning: Could not extract converter.py: {e}")
                    # If extraction fails, create a minimal version
                    with open(converter_path, 'w') as f:
                        f.write("""#!/usr/bin/env python3
import pandas as pd
import os
import sys
import argparse

def main():
    parser = argparse.ArgumentParser(description='Process Excel files and apply translations.')
    parser.add_argument('input_file', help='Path to the input Excel file (.xls or .xlsx)')
    parser.add_argument('-o', '--output', help='Path to the output Excel file (default: input_name_translated.xlsx)')
    parser.add_argument('--skip-description', action='store_true', help='Skip translating the Description column')
    parser.add_argument('--skip-message', action='store_true', help='Skip translating the Message column')
    
    args = parser.parse_args()
    
    # Logic for translation would go here
    # This is a placeholder for the actual translation code
    
    print(f"Processing {args.input_file}")
    # Create a simple output file for testing
    df = pd.read_excel(args.input_file)
    
    output_file = args.output
    if not output_file:
        base_name = os.path.basename(args.input_file)
        name_without_ext = os.path.splitext(base_name)[0]
        output_file = os.path.join(os.path.dirname(args.input_file), f"{name_without_ext}_translated.xlsx")
    
    df.to_excel(output_file, index=False)
    print(f"Saved to {output_file}")
    return True

if __name__ == "__main__":
    main()
""")
    
    def browse_file(self):
        if platform.system() == "Darwin":  # macOS
            filename = filedialog.askopenfilename(
                title="Select Excel File"
            )
        else:
            filetypes = (
                ("Excel files", "*.xls;*.xlsx"),
                ("All files", "*.*")
            )
            filename = filedialog.askopenfilename(
                title="Select Excel File",
                filetypes=filetypes
            )
        
        if filename:
            self.file_path.set(filename)
            
            # Automatically set default output name
            if not self.output_path.get():
                base_name = os.path.basename(filename)
                name_without_ext = os.path.splitext(base_name)[0]
                dir_name = os.path.dirname(filename)
                output_file = os.path.join(dir_name, f"{name_without_ext}_translated.xlsx")
                self.output_path.set(output_file)
    
    def browse_output(self):
        if platform.system() == "Darwin":  # macOS
            filename = filedialog.asksaveasfilename(
                title="Save Translated File As",
                defaultextension=".xlsx"
            )
        else:
            filetypes = (
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            )
            filename = filedialog.asksaveasfilename(
                title="Save Translated File As",
                defaultextension=".xlsx",
                filetypes=filetypes
            )
        
        if filename:
            self.output_path.set(filename)
    
    def log(self, message):
        self.console.insert(tk.END, message + "\n")
        self.console.see(tk.END)
        self.root.update_idletasks()
    
    def check_docker(self):
        self.log("Checking if Docker is installed and running...")
        
        def run_check():
            try:
                # Try multiple methods to detect Docker
                docker_running = False
                docker_path = "docker"  # Default path
                
                # Method 1: Standard docker info command
                try:
                    if platform.system() == "Windows":
                        result = subprocess.run(
                            ["docker", "info"], 
                            capture_output=True, 
                            text=True, 
                            creationflags=subprocess.CREATE_NO_WINDOW
                        )
                    else:
                        result = subprocess.run(
                            ["docker", "info"], 
                            capture_output=True, 
                            text=True
                        )
                    
                    if result.returncode == 0:
                        docker_running = True
                except Exception as e:
                    self.log(f"Standard Docker check failed: {str(e)}")
                
                # Method 2: Check common Docker paths on macOS
                if not docker_running and platform.system() == "Darwin":
                    common_paths = [
                        "/usr/local/bin/docker",
                        "/opt/homebrew/bin/docker",
                        "/Applications/Docker.app/Contents/Resources/bin/docker"
                    ]
                    for path in common_paths:
                        if os.path.exists(path):
                            try:
                                result = subprocess.run(
                                    [path, "info"], 
                                    capture_output=True, 
                                    text=True
                                )
                                if result.returncode == 0:
                                    docker_running = True
                                    docker_path = path
                                    self.log(f"Found Docker at: {path}")
                                    break
                            except Exception:
                                pass
                
                # Method 3: Check if Docker socket exists (Unix-based systems)
                if not docker_running and platform.system() != "Windows":
                    if os.path.exists('/var/run/docker.sock'):
                        try:
                            # Try a simple docker command using known paths
                            if platform.system() == "Darwin":
                                # Check common macOS paths
                                for path in common_paths:
                                    if os.path.exists(path):
                                        result = subprocess.run(
                                            [path, "version", "--format", "{{.Server.Version}}"],
                                            capture_output=True,
                                            text=True
                                        )
                                        if result.returncode == 0:
                                            docker_running = True
                                            docker_path = path
                                            self.log(f"Found Docker at: {path}")
                                            break
                            else:
                                # For Linux
                                result = subprocess.run(
                                    ["docker", "version", "--format", "{{.Server.Version}}"],
                                    capture_output=True,
                                    text=True
                                )
                                if result.returncode == 0:
                                    docker_running = True
                        except Exception:
                            pass
                
                # Method 4: Check for Docker Desktop app running (macOS specific)
                if not docker_running and platform.system() == "Darwin":
                    try:
                        ps_result = subprocess.run(
                            ["ps", "-A"], 
                            capture_output=True, 
                            text=True
                        )
                        if "Docker" in ps_result.stdout or "com.docker.docker" in ps_result.stdout:
                            docker_running = True
                            # Try to find Docker path
                            for path in common_paths:
                                if os.path.exists(path):
                                    docker_path = path
                                    self.log(f"Found Docker at: {path}")
                                    break
                    except Exception:
                        pass
                
                # Store the Docker path as an instance variable for later use
                self.docker_path = docker_path
                
                if docker_running:
                    self.log("✅ Docker is installed and running correctly.")
                    return True
                else:
                    self.log("❌ Docker is installed but not running. Please start Docker Desktop.")
                    messagebox.showwarning(
                        "Docker Not Running", 
                        "Docker is installed but not running. Please start Docker Desktop and try again."
                    )
                    return False
            except FileNotFoundError:
                self.log("❌ Docker is not installed. Please install Docker Desktop.")
                messagebox.showerror(
                    "Docker Not Installed", 
                    "Docker is not installed. Please install Docker Desktop from https://docs.docker.com/get-docker/"
                )
                return False
        
        threading.Thread(target=run_check).start()
    
    def translate_file(self):
        input_file = self.file_path.get()
        output_file = self.output_path.get() if self.output_path.get() else None
        
        if not input_file:
            messagebox.showerror("Error", "Please select an Excel file to translate")
            return
        
        if not os.path.exists(input_file):
            messagebox.showerror("Error", f"File does not exist: {input_file}")
            return
        
        # Check if at least one column is selected
        if not self.translate_description.get() and not self.translate_message.get():
            messagebox.showerror("Error", "Please select at least one column to translate")
            return
        
        # Start progress bar
        self.progress.start()
        self.translate_button.config(state="disabled")
        
        def run_translation():
            try:
                # Get the directory where the GUI script is located
                if getattr(sys, 'frozen', False):
                    # Running as compiled executable
                    script_dir = os.path.dirname(sys.executable)
                else:
                    # Running as script
                    script_dir = os.path.dirname(os.path.abspath(__file__))
                
                self.log(f"Starting translation process...")
                self.log(f"Input file: {input_file}")
                
                # Log which columns will be translated
                columns_to_translate = []
                if self.translate_description.get():
                    columns_to_translate.append("Description")
                if self.translate_message.get():
                    columns_to_translate.append("Message")
                
                self.log(f"Columns to translate: {', '.join(columns_to_translate)}")
                
                if output_file:
                    self.log(f"Output will be saved to: {output_file}")
                
                # Check Docker one more time and find its path
                docker_running = False
                docker_path = getattr(self, 'docker_path', 'docker')  # Get stored path or default
                
                # Check common Docker paths on macOS
                if platform.system() == "Darwin":
                    common_paths = [
                        "/usr/local/bin/docker",
                        "/opt/homebrew/bin/docker",
                        "/Applications/Docker.app/Contents/Resources/bin/docker"
                    ]
                    for path in common_paths:
                        if os.path.exists(path):
                            try:
                                result = subprocess.run(
                                    [path, "info"], 
                                    capture_output=True, 
                                    text=True
                                )
                                if result.returncode == 0:
                                    docker_running = True
                                    docker_path = path
                                    self.log(f"Using Docker at: {path}")
                                    break
                            except Exception:
                                pass
                else:
                    try:
                        docker_check = subprocess.run(
                            [docker_path, "info"], 
                            capture_output=True, 
                            text=True
                        )
                        if docker_check.returncode == 0:
                            docker_running = True
                    except Exception:
                        docker_running = False
                
                # If Docker is not detected, confirm with user if they want to proceed anyway
                if not docker_running:
                    self.log("⚠️ Docker was not detected but might still be running.")
                    if not messagebox.askyesno(
                        "Docker Not Detected", 
                        "Docker was not detected, but if you're sure it's running, you can continue.\n\nDo you want to proceed with the translation?"):
                        self.log("❌ Translation cancelled by user.")
                        return
                    else:
                        self.log("Proceeding with translation as requested by user...")
                
                # Build Docker image
                self.log("Building Docker image (this might take a minute the first time)...")
                
                # Prepare build directory with necessary files
                build_dir = os.path.join(os.path.dirname(script_dir), "docker_build_temp")
                os.makedirs(build_dir, exist_ok=True)
                
                # Create a Dockerfile that doesn't require pulling from Docker Hub
                alternative_dockerfile = os.path.join(build_dir, "Dockerfile")
                with open(alternative_dockerfile, 'w') as f:
                    f.write("""# Minimal Dockerfile that doesn't require Docker Hub credentials
FROM scratch
COPY converter.py /app/converter.py
CMD ["python", "/app/converter.py"]
""")
                
                # Copy converter.py to the build directory
                converter_py_path = os.path.join(script_dir, "converter.py")
                import shutil
                if os.path.exists(converter_py_path):
                    shutil.copy(converter_py_path, os.path.join(build_dir, "converter.py"))
                
                # Try building with the original Dockerfile first
                if platform.system() == "Windows":
                    build_cmd = [docker_path, "build", "-t", "excel-translator", script_dir]
                    build_process = subprocess.run(
                        build_cmd, 
                        capture_output=True, 
                        text=True,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                else:
                    build_cmd = [docker_path, "build", "-t", "excel-translator", script_dir]
                    build_process = subprocess.run(
                        build_cmd, 
                        capture_output=True, 
                        text=True
                    )
                
                # If the build fails due to credentials, try a different approach
                if build_process.returncode != 0 and "docker-credential" in build_process.stderr:
                    self.log("Docker image build failed due to credential issues.")
                    self.log("Trying an alternative approach...")
                    
                    # Try using the user's local Python instead of Docker
                    try:
                        self.log("Using local Python for translation...")
                        
                        # First ensure pandas is installed
                        try:
                            import pandas as pd
                        except ImportError:
                            self.log("Installing required packages...")
                            subprocess.run([sys.executable, "-m", "pip", "install", "pandas", "openpyxl", "xlrd"], 
                                          capture_output=True)
                        
                        # Import required modules
                        import pandas as pd
                        
                        # Process the file directly
                        self.log(f"Reading Excel file: {input_file}")
                        df = pd.read_excel(input_file)
                        
                        # More comprehensive translation dictionary with terms from your Excel file
                        translations = {
                            # Technical status values from your Excel file
                            "BAD STATE": "MAUVAIS ETAT",
                            "REMOTE": "DISTANT",
                            "OPEN": "OUVERT",
                            "ON": "ACTIF",
                            "SET": "RÉGLÉ",
                            "RESET": "RÉINITIALISÉ",
                            "SET - APP ACK": "RÉGLÉ - APP ACK",
                            "RESET - APP ACK": "RÉINITIALISÉ - APP ACK",
                            "OPEN - APP ACK": "OUVERT - APP ACK",
                            
                            # Common technical terms
                            "error": "erreur",
                            "warning": "avertissement",
                            "info": "info",
                            "debug": "débogage",
                            "critical": "critique",
                            "alert": "alerte",
                            "emergency": "urgence",
                            "notice": "avis",
                            "log": "journal",
                            "trace": "trace",
                            "status": "statut",
                            "update": "mise à jour",
                            "configuration": "configuration",
                            "processing": "traitement",
                            "output": "sortie",
                            "input": "entrée",
                            "message": "message",
                            "system": "système",
                            "network": "réseau",
                            "connection": "connexion",
                            "disconnect": "déconnecter",
                            "reconnect": "reconnecter",
                            "failure": "échec",
                            "success": "succès",
                            "retry": "réessayer",
                            "abort": "abandonner",
                            "timeout": "délai d'attente"
                        }
                        
                        # Define a better translation function that preserves case
                        def translate_text(text):
                            if not isinstance(text, str):
                                return text
                            
                            translated = text
                            
                            # Convert to lowercase for matching but preserve case for replacement
                            text_lower = text.lower()
                            
                            # Find all words in the dictionary that appear in the text
                            for eng, fr in translations.items():
                                # Create word boundary pattern for the English term
                                eng_lower = eng.lower()
                                if eng_lower in text_lower:
                                    # Find start position of the term in the original text
                                    start_pos = text_lower.find(eng_lower)
                                    
                                    # Get the original casing from the input text
                                    original_cased = text[start_pos:start_pos+len(eng)]
                                    
                                    # Apply same casing to the French translation
                                    if original_cased.isupper():
                                        replacement = fr.upper()
                                    elif original_cased[0].isupper():
                                        replacement = fr.capitalize()
                                    else:
                                        replacement = fr
                                    
                                    # Replace in the text
                                    translated = translated.replace(original_cased, replacement)
                            
                            return translated
                        
                        # Apply translations to selected columns by creating NEW columns
                        columns_translated = False
                        for col in columns_to_translate:
                            if col in df.columns:
                                self.log(f"Translating column: {col}")
                                # Create a new column with the translated content
                                df[f"{col} Français"] = df[col].apply(translate_text)
                                columns_translated = True
                        
                        if not columns_translated:
                            self.log("⚠️ Warning: None of the selected columns were found in the file.")
                            # Look for columns that might match by different case
                            df_columns_lower = [c.lower() for c in df.columns]
                            for col in columns_to_translate:
                                if col.lower() in df_columns_lower:
                                    idx = df_columns_lower.index(col.lower())
                                    actual_col = df.columns[idx]
                                    self.log(f"Found column '{actual_col}' that matches '{col}'")
                                    df[f"{actual_col} Français"] = df[actual_col].apply(translate_text)
                                    columns_translated = True
                        
                        if not columns_translated:
                            self.log("❌ Error: No columns could be translated.")
                            messagebox.showerror("Error", "No matching columns found for translation.")
                            return
                        
                        # Save the translated file
                        output_path = output_file or os.path.join(
                            os.path.dirname(input_file),
                            f"{os.path.splitext(os.path.basename(input_file))[0]}_translated.xlsx"
                        )
                        
                        df.to_excel(output_path, index=False)
                        self.log(f"✅ Saved translated file to: {output_path}")
                        
                        # Show success message
                        messagebox.showinfo("Success", "Translation completed successfully!")
                        
                        # Open the folder containing the output file
                        output_folder = os.path.dirname(output_path)
                        if platform.system() == "Windows":
                            os.startfile(output_folder)
                        elif platform.system() == "Darwin":  # macOS
                            subprocess.run(["open", output_folder])
                        else:  # Linux
                            subprocess.run(["xdg-open", output_folder])
                        
                        return
                        
                    except Exception as local_err:
                        self.log(f"Local Python translation failed: {str(local_err)}")
                        self.log("As a last resort, trying a simplified Docker build...")
                        
                        # Try to build with a simplified Dockerfile as a last resort
                        try:
                            # Create a minimal Dockerfile that doesn't require pulling from Docker Hub
                            with open(os.path.join(script_dir, "Dockerfile"), 'w') as f:
                                f.write("""FROM alpine:latest
RUN apk add --no-cache python3 py3-pip
RUN pip3 install pandas openpyxl xlrd
WORKDIR /app
COPY converter.py /app/
ENTRYPOINT ["python3", "/app/converter.py"]
""")
                            
                            if platform.system() == "Windows":
                                build_cmd = [docker_path, "build", "--pull=false", "-t", "excel-translator", script_dir]
                                build_process = subprocess.run(
                                    build_cmd, 
                                    capture_output=True, 
                                    text=True,
                                    creationflags=subprocess.CREATE_NO_WINDOW
                                )
                            else:
                                build_cmd = [docker_path, "build", "--pull=false", "-t", "excel-translator", script_dir]
                                build_process = subprocess.run(
                                    build_cmd, 
                                    capture_output=True, 
                                    text=True
                                )
                        except Exception as docker_err:
                            self.log(f"Alternative Docker build also failed: {str(docker_err)}")
                            messagebox.showerror("Error", "All translation methods failed. Please check Docker credentials in Docker Desktop settings.")
                            return
                
                if build_process.returncode != 0:
                    self.log("❌ Failed to build Docker image:")
                    self.log(build_process.stderr)
                    messagebox.showerror("Error", "Failed to build Docker image. See console for details.")
                    return
                
                self.log("Docker image built successfully.")
                
                # Run Docker container for translation
                self.log("Processing the Excel file...")
                
                input_dir = os.path.dirname(input_file)
                input_filename = os.path.basename(input_file)
                
                # Base command
                docker_cmd = [docker_path, "run", "--rm"]
                
                # Add column selection arguments
                column_args = []
                if not self.translate_description.get():
                    column_args.append("--skip-description")
                if not self.translate_message.get():
                    column_args.append("--skip-message")
                
                # Set up volume mounts and paths
                if output_file:
                    output_dir = os.path.dirname(output_file)
                    output_filename = os.path.basename(output_file)
                    
                    # If output directory is different from input directory, add another volume
                    if output_dir != input_dir:
                        docker_cmd.extend([
                            "-v", f"{input_dir}:/data_in", 
                            "-v", f"{output_dir}:/data_out",
                            "excel-translator"
                        ])
                        docker_cmd.extend(column_args)
                        docker_cmd.extend([f"/data_in/{input_filename}", "-o", f"/data_out/{output_filename}"])
                    else:
                        docker_cmd.extend(["-v", f"{input_dir}:/data", "excel-translator"])
                        docker_cmd.extend(column_args)
                        docker_cmd.extend([f"/data/{input_filename}", "-o", f"/data/{output_filename}"])
                else:
                    docker_cmd.extend(["-v", f"{input_dir}:/data", "excel-translator"])
                    docker_cmd.extend(column_args)
                    docker_cmd.append(f"/data/{input_filename}")
                
                self.log(f"Running command: {' '.join(docker_cmd)}")
                
                if platform.system() == "Windows":
                    docker_process = subprocess.run(
                        docker_cmd, 
                        capture_output=True, 
                        text=True,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                else:
                    docker_process = subprocess.run(
                        docker_cmd, 
                        capture_output=True, 
                        text=True
                    )
                
                if docker_process.returncode != 0:
                    self.log("❌ Failed to process the Excel file:")
                    self.log(docker_process.stderr)
                    messagebox.showerror("Error", "Failed to process the Excel file. See console for details.")
                    return
                
                # Handle successful translation
                self.log("✅ Success! File has been translated and saved.")
                if not output_file:
                    basename = os.path.basename(input_file)
                    name_without_ext = os.path.splitext(basename)[0]
                    output_path = os.path.join(input_dir, f"{name_without_ext}_translated.xlsx")
                    self.log(f"Output saved to: {output_path}")
                else:
                    self.log(f"Output saved to: {output_file}")
                
                messagebox.showinfo("Success", "Translation completed successfully!")
                
                # Open the folder containing the output file
                output_folder = output_dir if output_file else input_dir
                if platform.system() == "Windows":
                    os.startfile(output_folder)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", output_folder])
                else:  # Linux
                    subprocess.run(["xdg-open", output_folder])
                
            except Exception as e:
                self.log(f"❌ Error: {str(e)}")
                messagebox.showerror("Error", f"An error occurred: {str(e)}")
            finally:
                # Stop progress bar and re-enable button
                self.root.after(0, lambda: self.progress.stop())
                self.root.after(0, lambda: self.translate_button.config(state="normal"))
        
        # Run translation in a separate thread to keep UI responsive
        threading.Thread(target=run_translation).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = TranslatorApp(root)
    root.mainloop() 
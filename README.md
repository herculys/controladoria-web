# Controladoria Web Application

This project is a web interface for the Controladoria script, allowing users to select files for processing instead of hardcoding file paths in the script. The application is built using Flask and provides a user-friendly interface for file uploads.

## Project Structure

```
controladoria-web
├── src
│   ├── app.py                # Main entry point of the web application
│   ├── controladoria.py      # Original logic of the Controladoria script
│   ├── templates
│   │   └── index.html        # HTML template for file selection
│   └── static
│       └── style.css         # CSS styles for the web interface
├── requirements.txt          # List of dependencies
└── README.md                 # Project documentation
```

## Setup Instructions

1. **Clone the repository**:
   ```
   git clone <repository-url>
   cd controladoria-web
   ```

2. **Install dependencies**:
   It is recommended to use a virtual environment. You can create one using:
   ```
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```
   Then install the required packages:
   ```
   pip install -r requirements.txt
   ```

3. **Run the application**:
   Start the Flask web server by running:
   ```
   python src/app.py
   ```
   The application will be accessible at `http://127.0.0.1:5000`.

## Usage Guidelines

- Open your web browser and navigate to the provided URL.
- Use the file upload form to select the necessary files for processing.
- Follow the on-screen instructions to complete the process.

## Contributing

Contributions are welcome! Please feel free to submit a pull request or open an issue for any suggestions or improvements.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.
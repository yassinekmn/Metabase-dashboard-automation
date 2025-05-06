# Create and activate virtual environment (Windows)
python -m venv venv
.\venv\Scripts\activate

# Install requirements
pip install -r requirements.txt

# Install Playwright browsers
python -m playwright install chromium

# Create a .env File!!

# HOW TO CONFIGURE .env
# Metabase Credentials
METABASE_URL=https://analytics.example.com       # Your Metabase instance URL
METABASE_USERNAME=your_email@example.com        # Metabase login email
METABASE_PASSWORD=your_password_here            # Metabase password

# Email Settings (Gmail example)
SMTP_SERVER=smtp.gmail.com                      # SMTP server address
SMTP_PORT=587                                   # SMTP port (587 for TLS)
SENDER_EMAIL=your_email@gmail.com               # Email sender address
SENDER_PASSWORD=your_app_specific_password      # Gmail app password (recommended) NOT ACCOUNT PÄSSWORD!!!
RECIPIENTS=recipient1@example.com,recipient2@example.com  # Comma-separated recipients

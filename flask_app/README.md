# Translingoo Web Application

A Flask web application for translating Excel files from English to French, designed for deployment in an enterprise environment.

## Features

- Upload Excel files (.xls or .xlsx formats)
- Translate specific columns (Description and/or Message)
- Download translated Excel files
- Responsive web interface
- Handles large files up to 16MB

## Quick Deployment (For Testing)

The easiest way to deploy and test this application is using Docker:

### On macOS/Linux:

1. Make sure Docker is installed
2. Run the deployment script:
   ```bash
   ./deploy.sh
   ```

### On Windows:

1. Make sure Docker is installed
2. Run the deployment batch file:
   ```
   deploy.bat
   ```

The application will be available at http://127.0.0.1:5000/

## Enterprise Deployment (Recommended for 500+ Users)

For a company with 500+ users, we recommend deploying on an internal server or cloud infrastructure:

### Option 1: Internal Company Server (Recommended)

1. **Set up a Linux server** with Docker installed
2. Clone this repository to the server
3. Run the deployment script:
   ```bash
   cd flask_app
   ./deploy.sh
   ```
4. Configure Nginx as a reverse proxy (for production):

   ```nginx
   server {
       listen 80;
       server_name translingoo.example.com;

       location / {
           proxy_pass http://127.0.0.1:5000;
           proxy_set_header Host $host;
           proxy_set_header X-Real-IP $remote_addr;
       }
   }
   ```

5. Set up as a service:

   ```bash
   # Create a service file
   sudo nano /etc/systemd/system/translingoo.service

   # Add this content:
   [Unit]
   Description=Translingoo Excel Translator
   After=docker.service
   Requires=docker.service

   [Service]
   ExecStart=/bin/bash -c 'cd /path/to/Translingoo/flask_app && ./deploy.sh'
   Restart=always

   [Install]
   WantedBy=multi-user.target

   # Enable and start the service
   sudo systemctl enable translingoo
   sudo systemctl start translingoo
   ```

### Option 2: Cloud Deployment (e.g., Azure Container Instances)

1. Build the Docker image locally:

   ```bash
   cd flask_app
   ./deploy.sh
   ```

2. Tag and push to your container registry:

   ```bash
   docker tag translingoo mycompanyregistry.azurecr.io/translingoo:latest
   docker push mycompanyregistry.azurecr.io/translingoo:latest
   ```

3. Deploy to Azure Container Instances or similar service:
   - Set up with at least 2 CPU cores and 4GB RAM
   - Map port 5000
   - Set environment variable SECRET_KEY to a secure value

## Security Considerations for Enterprise Use

- Set a strong `SECRET_KEY` environment variable in production
- Implement user authentication with your company's SSO solution
- Store the application behind a company firewall or VPN
- Configure regular backups of the application data
- Set up monitoring for the application using your company's monitoring tools

## Usage Instructions for End Users

1. Open the application in your web browser (http://translingoo.example.com)
2. Select your Excel file (.xls or .xlsx format)
3. Choose which columns to translate
4. Click "Process File"
5. Download your translated file

## Support and Maintenance

- Regular updates should be scheduled monthly
- Monitor disk space on the server to ensure sufficient space for uploads
- Implement a cleanup routine to remove old files (> 30 days)

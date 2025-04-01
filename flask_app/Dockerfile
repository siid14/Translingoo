FROM python:3.9-slim

WORKDIR /app

# Copy requirements first to leverage Docker caching
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the current directory contents into the container
COPY . /app/

# Copy the parent src directory (needed for ExcelProcessor)
# This will be handled by a separate step during the build

# Set environment variables
ENV FLASK_APP=app.py
ENV PYTHONUNBUFFERED=1
ENV SECRET_KEY=change-this-in-production

# Create directories for uploads and downloads
RUN mkdir -p uploads downloads
RUN chmod 777 uploads downloads

# Expose port
EXPOSE 5000

# Run the application with Gunicorn
CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:5000", "app:app"] 
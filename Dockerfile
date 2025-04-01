FROM python:3.9-slim

WORKDIR /app

# Install required packages
RUN pip install pandas xlrd openpyxl odfpy

# Copy the converter script
COPY converter.py /app/

# Make the script executable
RUN chmod +x /app/converter.py

ENTRYPOINT ["python", "/app/converter.py"] 
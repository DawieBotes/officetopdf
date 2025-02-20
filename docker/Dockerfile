# Use the official Ubuntu Server image as the base
FROM ubuntu:latest

# Set environment variables to avoid interactive prompts
ENV DEBIAN_FRONTEND=noninteractive

# Update and install necessary packages
RUN apt-get update && apt-get upgrade -y && \
    apt-get install -y \
    curl \
    vim \
    libreoffice \
    python3 \
    python3-pip \
    python3-uno \
    netcat-openbsd \
    python3-flask \
    gunicorn \
    && rm -rf /var/lib/apt/lists/*

# Set the working directory inside the container
WORKDIR /app

# Copy the conversion script, API, and the start.sh script
COPY code/ /app/code/
COPY start.sh /app/start.sh

# Expose the default port for the API
EXPOSE 5000

# Make sure start.sh has execution permissions
RUN chmod +x /app/start.sh

# Set environment variables, including PYTHONPATH
ENV PYTHONPATH="/app"
ENV PORT=5000  
# Default port to 5000, can be overridden with ENV variable

# Start LibreOffice in headless mode and run the API using the start.sh script
CMD ["/app/start.sh"]

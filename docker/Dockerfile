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
    && rm -rf /var/lib/apt/lists/*

# Install Flask
RUN apt-get update && apt-get install -y python3-flask

# Set the working directory inside the container
WORKDIR /app

# Copy the conversion script and the API separately
COPY code/ /app/code/


# Expose port 5000 for the API
EXPOSE 5000

# Start LibreOffice in headless mode and run the API
CMD libreoffice --headless --accept="socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" & sleep 5 && python3 /app/code/api.py

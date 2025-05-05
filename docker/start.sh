#!/bin/bash

# Start LibreOffice in headless mode
libreoffice --headless --accept="socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" &

# Wait for LibreOffice to be ready
while ! nc -z localhost 2002; do
    sleep 1
done

# Get the port from the environment variable (default to 5000 if not set)
PORT=${PORT:-5000}

# Run the API using Gunicorn, bind to the specified port
exec gunicorn --timeout 300 -w 4 -b 0.0.0.0:$PORT code.api:app

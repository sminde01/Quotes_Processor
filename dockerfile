# layer 1: base image
FROM python:3.10-slim-bullseye

# layer 2: working directory
WORKDIR /app

# layer 3: system dependencies
RUN apt-get update && apt-get install -y 
    
# layer 4: copy app agnostic dependencies
COPY requirements.txt .

# layer 5: install python dependencies
RUN pip3 install -r requirements.txt

# layer 6: copy app source code (will be updated quite often)
COPY . .

# layer 7: update entrypoint.sh permissions
RUN chmod +x ./entrypoint.sh

# layer 8: setup listener port for the app
EXPOSE 8506

# layer 9: run entrypoint.sh script
ENTRYPOINT ["./entrypoint.sh"]
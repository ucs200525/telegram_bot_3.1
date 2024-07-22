# Use an official Python runtime as a parent image
FROM python:3.9-slim

# Set the working directory in the container
WORKDIR /usr/src/app

# Copy the rest of the Node.js application code to the working directory
COPY . .

# Set the working directory in the container
WORKDIR /usr/src/app

# Copy installed Node.js modules from the previous stage
COPY --from=node-builder /usr/src/app /usr/src/app

# Create a virtual environment for Python
RUN python -m venv /venv

# Activate the virtual environment and install Python requirements
COPY requirements.txt /app/
RUN /venv/bin/pip install --no-cache-dir -r /app/requirements.txt

# Copy all other application files
COPY . /app/

# Set the virtual environment's bin directory in PATH
ENV PATH="/venv/bin:$PATH"

# Expose the port the app runs on
EXPOSE 8080

# Run Python application
CMD ["python", "bot.py"]

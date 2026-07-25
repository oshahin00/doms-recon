# 1. Start with a lightweight Linux image that has Python pre-installed
FROM python:3.9-slim

# 2. Set the working directory in the container
WORKDIR /app

# 3. Copy only the requirements file first (this caches the installation step to make future builds faster)
COPY requirements.txt .

# 4. Install the Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# 5. Copy the rest of the application code into the container
COPY . .

# 6. Expose the port that the application will run on
EXPOSE 8501

# 7. Tell the container how to run the application
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
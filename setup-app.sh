#!/bin/bash

echo "1. Updating the Linux system..."
apt-get update -y

echo "2. Installing Docker and Git..."
apt-get install -y docker.io git

echo "3. Starting the Docker engine..."
systemctl start docker
systemctl enable docker

echo "4. Downloading your code from GitHub..."
# Clean the directory first to prevent conflicts
rm -rf /opt/doms-recon  

# Notice the /opt/doms-recon at the very end of this line! That was missing.
git clone https://github.com/oshahin00/doms-recon.git /opt/doms-recon

echo "5. Building the Docker container..."
# Now this cd command will work because the folder actually exists
cd /opt/doms-recon
docker build -t doms-app .

echo "6. Running the application..."
docker rm -f doms-container 2>/dev/null || true

docker run -d \
  --restart unless-stopped \
  --name doms-container \
  -v /opt/doms-recon:/app \
  -p 8501:8501 \
  doms-app \
  streamlit run app.py --server.address=0.0.0.0 --server.enableCORS=false --server.enableXsrfProtection=false

echo "✅ App Server Provisioning Complete!"
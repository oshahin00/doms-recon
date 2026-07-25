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


echo "7. Setting up Server Monitoring (Node Exporter, Prometheus, Grafana)..." 

# Create a dedicated directory for monitoring configs
mkdir -p /opt/monitoring/prometheus

# Automatically generate the Prometheus config file
cat <<EOF > /opt/monitoring/prometheus/prometheus.yml
global: 
    scrape_interval: 15s

scrape_configs:
    - job_name: 'appserver-hardware'
      static_configs:
          - targets: ['192.168.50.11:9100']

EOF

# Remove exising monitoring containers if they exist (ensure the script can be run multiple times safely)
docker rm -f node-exporter prometheus grafana 2>/dev/null || true

# Boot Node Exporter (The Hardware Agent)
docker run -d \
    --name node-exporter \
    --restart unless-stopped \
    --net="host" \
    --pid="host" \
    -v "/:/host:ro,rslave" \
    -p 9100:9100 \
    prom/node-exporter

# Boot Prometheus (The Database)
docker run -d \
    --name prometheus \
    --restart unless-stopped \
    -p 9090:9090 \
    -v /opt/monitoring/prometheus/prometheus.yml:/etc/prometheus/prometheus.yml \
    prom/prometheus

# Boot Grafana (The Dashboard)
docker run -d \
    --name grafana \
    --restart unless-stopped \
    -p 3000:3000 \
    -e "GF_INSTALL_PLUGINS=frser-sqlite-datasource" \
    -v grafana-storage:/var/lib/grafana \
    -v /opt/doms-recon:/app-data:ro \
    grafana/grafana

echo "Monitoring Stack Deployed Successfully!"
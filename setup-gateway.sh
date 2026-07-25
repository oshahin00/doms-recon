#!/bin/bash 

echo "1. Updating system and installing Nginx..." 
apt-get update -y 
apt-get install -y nginx 

echo "2. Configuring Nginx as a Reverse Proxy for Streamlit..." 
cat <<EOF > /etc/nginx/sites-available/default
server {
    listen 80;
    server_name _;

    location / {
        proxy_pass http://192.168.50.11:8501/;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        
        # Required for Streamlit WebScockets
        proxy_http_version 1.1;
        proxy_set_header Upgrade \#http_upgarade;
        proxy_set_header Connection "upgrade";
    }
    # Dedicated block strictly for Streamlit WebSockets
    location ^~ /_stcore/stream {
        proxy_pass http://192.168.50.11:8501/_stcore/stream;
        proxy_http_version 1.1;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header Host \$host;
        proxy_set_header Upgrade \$http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_read_timeout 86400;
    }
}
EOF

echo "3. Restarting Nginx to apply changes..."
systemctl restart nginx
systemctl enable nginx

echo "✅ Gateway Provisioning Complete!"
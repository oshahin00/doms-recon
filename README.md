# doms-recon: Automated Deployment Pipeline[cite: 5]

This environment utilizes Infrastructure as Code (IaC) to automatically provision a secure, two-tier network architecture[cite: 5]. By separating public-facing traffic routing from backend application processing, it successfully mimics a secure enterprise deployment where backend systems are managed remotely[cite: 5].

## 🏗️ Architecture Overview

The infrastructure relies on virtualization to maintain a consistent environment across deployments.

| Component | Details | Service / Role |
| :--- | :--- | :--- |
| **Host Machine** | Developer Workstation[cite: 5] | Bare-metal execution environment[cite: 5] |
| **Hypervisor** | VirtualBox (Managed via Vagrant)[cite: 5] | Virtualization layer[cite: 5] |
| **Node 1 (Gateway)** | Ubuntu 22.04 LTS (`192.168.50.10`)[cite: 5] | Nginx Reverse Proxy[cite: 5] |
| **Node 2 (App Server)** | Ubuntu 22.04 LTS (`192.168.50.11`)[cite: 5] | Docker & Streamlit[cite: 5] |

## 📦 Application Payload

The core workload is a custom Python application named `doms-recon`[cite: 5]. 

*   **Containerization:** Packaged via Docker to ensure consistent execution regardless of the underlying host operating system[cite: 5].
*   **Frontend Interface:** Utilizes Streamlit to serve the Python backend logic as an interactive, real-time web dashboard[cite: 5].
*   **Networking:** The application is exposed internally on port `8501`[cite: 5].

## ⚙️ Infrastructure & Provisioning

The environment is entirely automated using Vagrant and Bash scripting, eliminating the need for manual server configuration[cite: 5].

### Configuration Files

*   **`Vagrantfile`**: Defines the hardware specifications (RAM/CPU), assigns static IPs on a private internal network, and maps port `8080` on the host machine to port `80` on the Gateway[cite: 5].
*   **`setup-app.sh`**: A Bash script executed automatically on the App Server upon boot[cite: 5]. It updates the operating system, installs Docker and Git, pulls the latest application code from GitHub, and builds/runs the Docker container in detached mode with auto-restart policies[cite: 5].
*   **`setup-gateway.sh`**: A Bash script executed on the Gateway node[cite: 5]. It installs Nginx and configures a reverse proxy block to securely bridge external HTTP traffic and WebSockets from port `80` to the hidden App Server's port `8501`[cite: 5].

## 🐛 Troubleshooting Ledger

During the initial build phase, several environmental and routing anomalies were encountered and resolved[cite: 5].

| Symptom / Error | Root Cause | Resolution |
| :--- | :--- | :--- |
| **Missing Directory Error during Docker Build** | The `git clone` command downloaded the repository into a default folder instead of the specific `/opt/` path expected by the subsequent Docker build command[cite: 5]. | Appended the explicit target directory path (`/opt/doms-recon`) to the end of the `git clone` command in the Bash script[cite: 5]. |
| **Browser `ERR_CONNECTION_TIMED_OUT`** | The host machine's network adapter failed to route browser traffic directly to the VirtualBox Host-Only private IP subnet (`192.168.50.11`)[cite: 5]. | Implemented a Gateway VM using Vagrant port forwarding (`localhost:8080` -> `guest:80`) to bypass Windows routing limitations[cite: 5]. |
| **`502 Bad Gateway` (Nginx)** | VirtualBox failed to assign the static IP (`192.168.50.11`) to the App Server's network interface during the heavy initial boot/provisioning cycle[cite: 5]. | Executed `vagrant reload appserver` to force hardware re-initialization, then manually ran `docker start doms-container` to wake the sleeping application[cite: 5]. |
| **Blank Gray Screen & WebSocket Failure** | Streamlit's built-in CORS and Cross-Site Request Forgery (XSRF) security protocols actively blocked the connection because the traffic originated from a proxy (`localhost:8080`) instead of its native IP[cite: 5]. | Appended `--server.enableCORS=false` and `--server.enableXsrfProtection=false` flags to the `docker run` command to authorize the Nginx proxy traffic[cite: 5]. |
| **Persistent WebSocket Disconnect** | The Nginx configuration included a trailing slash in the proxy pass URL, which stripped the routing structure required by Streamlit's `/_stcore/stream` endpoint[cite: 5]. | Removed the trailing slash and implemented a dedicated `location ^~ /_stcore/stream` block in Nginx with explicit `Upgrade` headers for WebSocket traffic[cite: 5]. |

> **TODO:** Add repository cloning instructions, environment variable setup, and local development usage examples.
Vagrant.configure("2") do |config|
  # Increase global timeouts to reduce "timeout" errors during `vagrant up`
  # We will use Ubuntu as our base Linux OS
  config.vm.box = "ubuntu/jammy64"

  # ---------------------------------------------------
  # VM 1: The App Server (Where Docker will run)
  # ---------------------------------------------------
  config.vm.define "appserver" do |app|
    app.vm.hostname = "appserver"
    # Create a private network
    app.vm.network "private_network", ip: "192.168.50.11"
    
    # NEW LINE: Tell Vagrant to run our bash script when booting up!
    app.vm.provision "shell", path: "app.sh"

    # Hardware Allocation: 2GB RAM, 2 CPUs
    app.vm.provider "virtualbox" do |vb|
      vb.name = "appserver"
      vb.memory = "2048"
      vb.cpus = 2
    end
  end

  # ---------------------------------------------------
  # VM 2: The Gateway (Reverse Proxy)
  # ---------------------------------------------------
  config.vm.define "gateway" do |gw|
    gw.vm.hostname = "gateway"
    # Put it on the same private network as the app server
    gw.vm.network "private_network", ip: "192.168.50.10"
    
    # Port Forwarding: Access the gateway from your laptop via port 8080
    gw.vm.network "forwarded_port", guest: 80, host: 8080

    # NEW LINE: Tell Vagrant to run our gateway script when booting up!
    gw.vm.provision "shell", path: "setup-gateway.sh"
    
    # Hardware Allocation: 1GB RAM, 1 CPU
    gw.vm.provider "virtualbox" do |vb|
      vb.name = "gateway"
      vb.memory = "2048"
      vb.cpus = 1
    end
  end
end


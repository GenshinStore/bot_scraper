# Update sistem dan install pip
sudo apt update
sudo apt install python3-pip python3-venv -y

# Masuk ke folder bot Anda
cd /path/ke/folder/bot/anda

# Buat virtual environment (direkomendasikan)
python3 -m venv venv
source venv/bin/activate

# Instal requirement
pip install -r requirements.txt
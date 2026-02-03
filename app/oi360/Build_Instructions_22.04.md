# 🚀 Building Oi360 SOA RECO on Ubuntu 22.04

## 📦 Step 1: Copy to Flash Drive

Create folder `Oi360_SOA_Source` and copy these files:

### ✅ Files to Copy

- `oi_360_soa_reco_pyqt_final.py`
- `Oi360 Logo_4.png`
- `logo.png`
- `requirements.txt`
- `install.sh`
- `uninstall.sh`

### ❌ DO NOT Copy

- `venv/`
- `__pycache__/`
- `.git/`

---

## 🛠️ Step 2: On Ubuntu 22.04

Open terminal in the folder and run:

```bash
# Install system requirements
sudo apt update
sudo apt install -y python3-pip python3-venv python3-dev build-essential libgl1-mesa-dev libxkbcommon-x11-0 libfuse2

# Create virtual environment
python3 -m venv venv
source venv/bin/activate

# Install Python packages
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller
```

---

## 🏗️ Step 3: Build

```bash
pyinstaller --noconfirm --onefile --windowed --name "Oi360_SOA_RECO" \
 --add-data "Oi360 Logo_4.png:." \
 --hidden-import "PyQt5" \
 --hidden-import "pandas" \
 --hidden-import "openpyxl" \
 --hidden-import "xlrd" \
 --hidden-import "xlsxwriter" \
 oi_360_soa_reco_pyqt_final.py
```

---

## 🚀 Step 4: Deploy

1. Test: `./dist/Oi360_SOA_RECO`
2. Copy to other machines: `dist/Oi360_SOA_RECO` + `logo.png` + `install.sh`
3. Run `./install.sh` on each machine

Regarding Logo issue

The gear wheel icon means Ubuntu can't find the logo file. This is usually because:

The icon wasn't copied correctly to ~/.local/share/icons/hicolor/256x256/apps/
The icon cache wasn't refreshed
Run these commands on your Ubuntu 22.04 machine to fix it:

bash

# 1. Check if the icon file exists

ls -la ~/.local/share/icons/hicolor/256x256/apps/

# 2. If it's missing, copy it manually (from your dist folder)

cp /path/to/your/dist/logo.png ~/.local/share/icons/hicolor/256x256/apps/oi360-pdf-suite.png

# 3. Force refresh the icon cache

gtk-update-icon-cache -f ~/.local/share/icons/hicolor/

# 4. Restart GNOME shell (press Alt+F2, type 'r', press Enter)

# OR log out and log back in

Also check: Is the logo.png file actually on your dist folder on the 22.04 machine? Make sure you copied it along with
install.sh
 and Oi360_Suite.

If the issue persists, try using the absolute path method. Run this on 22.04:

bash

# Edit the desktop file to use absolute path for icon

sed -i "s|Icon=oi360-pdf-suite|Icon=$HOME/.local/share/icons/hicolor/256x256/apps/oi360-pdf-suite.png|" ~/.local/share/applications/oi360-pdf-suite.desktop
Then log out and back in to refresh the desktop.

#!/bin/bash
# Installatie voor marktplaats_productmanager.py
# Draai dit vanuit dezelfde map als marktplaats_manager.py en auto_marktplaats.py

set -e

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "📦 GTK/PyGObject en Python-afhankelijkheden installeren..."
sudo apt update
sudo apt install -y python3-gi gir1.2-gtk-3.0 python3-pip

echo "📦 Python-packages installeren (barcode, sheets-ondersteuning)..."
pip install --break-system-packages python-barcode pillow

echo "📦 Printopties uitbreiden..."
# Voor Excel export
pip install openpyxl --break-system-packages

# Voor barcode generatie (al vereist)
pip install python-barcode pillow --break-system-packages

# Voor brother_ql printer
pip install brother_ql --break-system-packages

# Excel export
pip install openpyxl --break-system-packages

# Voor prullenbak (verwijderen naar prullenbak)
pip install Send2Trash --break-system-packages

# Icoon plaatsen (als icon.png naast dit script staat)
if [ -f "$APP_DIR/icon.png" ]; then
    echo "🖼️  Icoon gevonden, wordt gebruikt voor de snelkoppeling"
else
    echo "⚠️  Geen icon.png gevonden naast install_productmanager.sh - snelkoppeling krijgt geen custom icoon"
fi

# Startscript aanmaken
cat > "$APP_DIR/start_productmanager.sh" << EOF
#!/bin/bash
cd "\$(dirname "\$0")"
python3 marktplaats_productmanager.py
EOF
chmod +x "$APP_DIR/start_productmanager.sh"
echo "✅ start_productmanager.sh aangemaakt"

# Desktop-snelkoppeling aanmaken
cat > "$APP_DIR/MarktplaatsProductManager.desktop" << EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=Marktplaats Product Manager
Comment=Registreer en beheer producten voor Marktplaats
Exec=$APP_DIR/start_productmanager.sh
Icon=$APP_DIR/icon.png
Terminal=false
StartupNotify=true
StartupWMClass=marktplaats_productmanager
Categories=Utility;Office;
EOF
chmod +x "$APP_DIR/MarktplaatsProductManager.desktop"

# Ook naar het applicatiemenu kopiëren, zodat 'ie doorzoekbaar is
mkdir -p ~/.local/share/applications
cp "$APP_DIR/MarktplaatsProductManager.desktop" ~/.local/share/applications/
update-desktop-database ~/.local/share/applications 2>/dev/null || true
echo "✅ Snelkoppeling aangemaakt en toegevoegd aan het applicatiemenu"

echo ""
echo "✅ Klaar. Start de app op 2 manieren:"
echo "   1. cd $APP_DIR && ./start_productmanager.sh"
echo "   2. Zoek 'Marktplaats Product Manager' in het applicatiemenu"
echo ""
echo "📌 Open daarna meteen '⚙️ Instellingen' om:"
echo "   - de opslagmethode te kiezen (XML of Google Sheets)"
echo "   - de basismap in te stellen (moet gelijk zijn aan marktplaats_manager.py)"
echo "   - opslaglocatie/sublocatie/rij-codes in te vullen"
echo "   - eventueel een labelprinter te koppelen"

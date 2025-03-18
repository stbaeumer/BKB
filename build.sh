#!/bin/bash

# Version und Zeitstempel für den Namen
#VERSION="v0.0.1"
#TIMESTAMP=$(date +%Y%m%d%H%M%S)
#OUTPUT_DIR="/workspaces/BKB/publish"
#OUTPUT_DIR="/home/stefan/RiderProjects/BKB/publish" 


# Erstelle das Veröffentlichungsverzeichnis (falls nicht vorhanden)
#mkdir -p "$OUTPUT_DIR"

#cd /workspaces/BKB/gui

#/home/stefan/.dotnet/dotnet publish -c Release -r linux-x64 --self-contained true -p:PublishSingleFile=true -o "$OUTPUT_DIR"
#mv "$OUTPUT_DIR/gui" "$OUTPUT_DIR/BKB-${VERSION}-Linux"
#chmod +x "$OUTPUT_DIR/BKB-${VERSION}-Linux"

#cd /home/stefan/RiderProjects/BKB/shell
#/home/stefan/.dotnet/dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o "$OUTPUT_DIR"
#mv "$OUTPUT_DIR/gui.exe" "$OUTPUT_DIR/BKB.exe"
#cp "$OUTPUT_DIR/BKB.exe" /home/stefan/Windows/SchILD-NRW/

cd /home/stefan/RiderProjects/BKB/shell
/home/stefan/.dotnet/dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
cp /home/stefan/RiderProjects/BKB/shell/bin/Release/net8.0/win-x64/publish/* /home/stefan/Windows/SchILD-NRW/publish/
#/home/stefan/.dotnet/dotnet publish -c Release -r linux-x64 --self-contained true -p:PublishSingleFile=true -o "$OUTPUT_DIR"

#mv "$OUTPUT_DIR/shell.exe" "$OUTPUT_DIR/BKB-Shell-${VERSION}-Windows.exe"
#mv "$OUTPUT_DIR/shell" "$OUTPUT_DIR/BKB-Shell-${VERSION}-Linux"
#chmod +x "$OUTPUT_DIR/BKB-Shell-${VERSION}-Linux"


echo "Veröffentlichung abgeschlossen!"


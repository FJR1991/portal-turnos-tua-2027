#!/bin/sh
# Marca una versión nueva del portal: así los navegadores que tengan la
# anterior en caché se recargan solos al abrirlo. Ejecutar antes de cada commit.
cd "$(dirname "$0")/.." || exit 1
B=$(date +%Y%m%d%H%M)
sed -i.bak -E "s/var BUILD=\"[0-9]+\"/var BUILD=\"$B\"/" index.html && rm -f index.html.bak
printf '%s\n' "$B" > version.txt
echo "Versión $B"

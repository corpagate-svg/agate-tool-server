#!/usr/bin/env bash
set -euo pipefail

if [ "$(id -u)" -ne 0 ]; then
  echo "root で実行してください(例: sudo bash install.sh)" >&2
  exit 1
fi

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ENV_FILE="$APP_DIR/server/.env"

echo "=== 1/6: Node.js の確認 ==="
if ! command -v node >/dev/null 2>&1; then
  echo "Node.js が見つかりません。先に Node.js をインストールしてください。" >&2
  exit 1
fi
node -v

echo "=== 2/6: サーバープログラムの依存パッケージをインストール ==="
cd "$APP_DIR/server"
npm install --omit=dev

echo "=== 3/6: 設定ファイル(.env)の準備 ==="
if [ -f "$ENV_FILE" ]; then
  echo ".env は既にあるのでそのまま使います(トークンは変わりません)"
  # shellcheck disable=SC1090
  source "$ENV_FILE"
else
  API_TOKEN="$(openssl rand -hex 24)"
  {
    echo "PORT=3000"
    echo "API_TOKEN=$API_TOKEN"
  } > "$ENV_FILE"
  chmod 600 "$ENV_FILE"
fi

echo "=== 4/6: HTTPS用の住所(ドメイン)を決定 ==="
PUBLIC_IP="$(curl -fsSL https://api.ipify.org || hostname -I | awk '{print $1}')"
DOMAIN="$(echo "$PUBLIC_IP" | tr '.' '-').nip.io"
echo "ドメイン: $DOMAIN"

echo "=== 5/6: Caddy(HTTPS化)のインストールと設定 ==="
if ! command -v caddy >/dev/null 2>&1; then
  apt install -y debian-keyring debian-archive-keyring apt-transport-https curl gnupg
  curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/gpg.key' | gpg --dearmor -o /usr/share/keyrings/caddy-stable-archive-keyring.gpg
  curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/debian.deb.txt' | tee /etc/apt/sources.list.d/caddy-stable.list
  apt update
  apt install -y caddy
fi
sed "s/__DOMAIN__/$DOMAIN/" "$APP_DIR/Caddyfile.template" > /etc/caddy/Caddyfile
systemctl enable caddy >/dev/null 2>&1 || true
systemctl restart caddy

echo "=== 6/6: サーバープログラムを自動起動に登録 ==="
sed "s#__APP_DIR__#$APP_DIR#g" "$APP_DIR/systemd/agate-tool-server.service.template" > /etc/systemd/system/agate-tool-server.service
systemctl daemon-reload
systemctl enable agate-tool-server >/dev/null 2>&1 || true
systemctl restart agate-tool-server

if command -v ufw >/dev/null 2>&1; then
  ufw allow 80/tcp >/dev/null 2>&1 || true
  ufw allow 443/tcp >/dev/null 2>&1 || true
fi

sleep 2

# shellcheck disable=SC1090
source "$ENV_FILE"
echo ""
echo "======================================"
echo " セットアップ完了"
echo "======================================"
echo "URL     : https://$DOMAIN"
echo "トークン: $API_TOKEN"
echo ""
echo "この2つを控えて、Claudeとの会話に伝えてください。"
echo "トークンは合言葉のようなものなので、他の人に教えないでください。"
echo "======================================"

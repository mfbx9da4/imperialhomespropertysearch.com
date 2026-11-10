#!/bin/sh
# Local testing only: saves the BT email password into mailer/.dev.vars (ignored by git),
# so `npx wrangler dev` can send through BT. Typing is hidden and the password is never printed.
set -eu
cd "$(dirname "$0")"

trap 'stty echo 2>/dev/null' EXIT INT TERM
printf 'BT email password for imperialhomes@btinternet.com (typing is hidden): '
stty -echo
IFS= read -r pass
stty echo
printf '\n'

if [ -z "$pass" ]; then
  echo "Nothing typed. Nothing saved."
  exit 1
fi
case "$pass" in
  *"'"*)
    echo "The password has a ' character, which this script cannot save. Edit mailer/.dev.vars by hand."
    exit 1
    ;;
esac

touch .dev.vars
chmod 600 .dev.vars
grep -v '^SMTP_PASS=' .dev.vars > .dev.vars.tmp || true
printf "SMTP_PASS='%s'\n" "$pass" >> .dev.vars.tmp
mv .dev.vars.tmp .dev.vars
chmod 600 .dev.vars
pass=''
echo "Saved to mailer/.dev.vars. You can close this tab."

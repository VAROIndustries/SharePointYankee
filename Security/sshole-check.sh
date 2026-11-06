#!/bin/sh
# sshole-check.sh
# Read-only check for *sshole (CVE-2026-123158): OpenSSH on musl libc without PAM,
# where a 257 to 1024 byte password can log in to an account whose shadow field is "*".
# Checks the host and every running Docker container. Changes nothing, prints no hashes.
#
# Blog Post: https://sharepointyankee.com/2026/11/06/sshole-checking-your-linux-hosts-and-docker-containers-for-the-openssh-musl-password-bypass/
#
# Usage (run as root, or it can't read shadow files or query Docker):
#   sudo sh sshole-check.sh
#   ssh root@myserver 'sh -s' < sshole-check.sh
#
# Output:
#   FLAG   sshd on musl. Compare pkg to your distro's fixed version, then patch.
#   ok     sshd on glibc, not affected by this bug.
#   CHECK  sshd signs in a container with no shell, so look at it by hand.
#   star=  accounts with "*" in shadow    hash=  accounts with a $5$ or $6$ hash
#   Exploitable needs: FLAG, passwordauthentication yes, no "usepam yes", star>0, hash>0.
#   A build without PAM does not print usepam at all. Empty cfg means sshd -T could not run.
#
# Exit code: 0 nothing flagged, 1 something flagged, 2 could not check everything.

# Runs on the host and inside each container. Prints: musl|sshd|star|hash|pkg|cfg
PROBE='
musl=$(ls /lib/ld-musl-* 2>/dev/null | wc -l | tr -d " ")
bin=$(command -v sshd 2>/dev/null || ls /usr/sbin/sshd 2>/dev/null)
pkg=$(apk info -v 2>/dev/null | grep "^openssh-server" | tr "\n" " ")
cfg=
[ -n "$bin" ] && cfg=$("$bin" -T 2>/dev/null | grep -Ei "^(usepam|passwordauthentication) " | tr "\n" " ")
star=$(awk -F: "\$2 == \"*\" { n++ } END { print n+0 }" /etc/shadow 2>/dev/null)
hash=$(awk -F: "\$2 ~ /^[\$][56][\$]/ { n++ } END { print n+0 }" /etc/shadow 2>/dev/null)
echo "$musl|${bin:--}|${star:-?}|${hash:-?}|$pkg|$cfg"
'

flagged=0
incomplete=0

# report <name> <probe output> <running count or -> <published port 22>
report() {
  IFS='|' read -r musl bin star hash pkg cfg <<EOF
$2
EOF
  if [ "$bin" = "-" ] && [ "$3" = "0" -o "$3" = "-" ] && [ -z "$4" ]; then
    return 1
  fi
  detail="sshd=$bin running=$3 port22=[$4] pkg=[$pkg] cfg=[$cfg] star=$star hash=$hash"
  if [ "$musl" -gt 0 ]; then
    flagged=$((flagged + 1))
    echo "FLAG  $1 musl $detail"
  else
    echo "ok    $1 glibc $detail"
  fi
}

host=$(sh -c "$PROBE" </dev/null)
echo "HOST $(hostname) os=$([ -r /etc/os-release ] && . /etc/os-release; echo "${PRETTY_NAME:-unknown}")"
report "$(hostname)" "$host" "-" "" || echo "none  $(hostname) no sshd found"

total=0
no_shell=0
if command -v docker >/dev/null 2>&1; then
  if names=$(docker ps --format '{{.Names}}' 2>/dev/null); then
    for n in $names; do
      total=$((total + 1))
      # docker top and docker port work even on images without a shell
      running=$(docker top "$n" 2>/dev/null | grep -Ec '[ /]sshd( |:|$)')
      port22=$(docker port "$n" 22/tcp 2>/dev/null | tr '\n' ' ')
      if ! docker exec "$n" sh -c true >/dev/null 2>&1; then
        no_shell=$((no_shell + 1))
        if [ "$running" -gt 0 ] || [ -n "$port22" ]; then
          incomplete=1
          echo "CHECK $n no shell, running=$running port22=[$port22]"
        fi
        continue
      fi
      report "$n" "$(docker exec "$n" sh -c "$PROBE" </dev/null)" "$running" "$port22"
    done
  else
    incomplete=1
    echo "WARN  docker is installed but containers could not be listed (run as root?)"
  fi
else
  echo "none  docker not installed, host check only"
fi

echo "SUMMARY containers=$total flagged=$flagged no_shell_uninspected=$no_shell incomplete=$incomplete"
[ "$flagged" -gt 0 ] && exit 1
[ "$incomplete" -gt 0 ] && exit 2
exit 0

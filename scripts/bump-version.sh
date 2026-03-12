#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
version_file="${repo_root}/VERSION"

if [[ ! -f "${version_file}" ]]; then
  printf "1.00\n" > "${version_file}"
fi

current="$(tr -d '[:space:]' < "${version_file}")"
if [[ ! "${current}" =~ ^[0-9]+\.[0-9]{2}$ ]]; then
  echo "Invalid VERSION format: ${current}. Expected N.NN" >&2
  exit 1
fi

major="${current%.*}"
minor="${current#*.}"
current_cents=$((10#${major} * 100 + 10#${minor}))
next_cents=$((current_cents + 1))
next_major=$((next_cents / 100))
next_minor=$((next_cents % 100))
next_version="$(printf "%d.%02d" "${next_major}" "${next_minor}")"

printf "%s\n" "${next_version}" > "${version_file}"
printf "window.APP_VERSION = \"%s\";\n" "${next_version}" > "${repo_root}/web/version.js"
printf "window.APP_VERSION = \"%s\";\n" "${next_version}" > "${repo_root}/docs/version.js"

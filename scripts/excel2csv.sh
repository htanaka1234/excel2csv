#!/usr/bin/env bash
set -euo pipefail

IMAGE="${EXCEL2CSV_IMAGE:-excel2csv:local}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

usage() {
  cat <<'USAGE'
Usage:
  scripts/excel2csv.sh [--gzip] [-o OUTPUT] [--password PASSWORD] INPUT...

Examples:
  scripts/excel2csv.sh -o out/merged.csv data/a.xlsx data/b.xlsx
  scripts/excel2csv.sh --gzip -o out/merged.csv.gz data/
USAGE
}

gzip_output=0
output=""
password="${EXCEL2CSV_PASSWORD:-}"
inputs=()

while [[ $# -gt 0 ]]; do
  case "$1" in
    --gzip)
      gzip_output=1
      shift
      ;;
    -o|--output)
      output="${2:-}"
      shift 2
      ;;
    --password)
      password="${2:-}"
      shift 2
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    --)
      shift
      while [[ $# -gt 0 ]]; do
        inputs+=("$1")
        shift
      done
      ;;
    *)
      inputs+=("$1")
      shift
      ;;
  esac
done

if [[ ${#inputs[@]} -eq 0 ]]; then
  usage >&2
  exit 2
fi

if [[ -z "${output}" ]]; then
  first_input="$(realpath "${inputs[0]}")"
  if [[ -d "${first_input}" ]]; then
    output_dir="${first_input}"
  else
    output_dir="$(dirname "${first_input}")"
  fi
  stamp="$(date +%Y%m%d_%H%M%S)"
  if [[ "${gzip_output}" -eq 1 ]]; then
    output="${output_dir}/merged_${stamp}.csv.gz"
  else
    output="${output_dir}/merged_${stamp}.csv"
  fi
fi

output_abs="$(realpath -m "${output}")"
output_dir="$(dirname "${output_abs}")"
output_name="$(basename "${output_abs}")"
mkdir -p "${output_dir}"

if ! docker image inspect "${IMAGE}" >/dev/null 2>&1; then
  docker build -t "${IMAGE}" "${PROJECT_ROOT}"
fi

docker_args=(run --rm -v "${output_dir}:/output")
container_inputs=()

for index in "${!inputs[@]}"; do
  input_abs="$(realpath "${inputs[$index]}")"
  input_name="$(basename "${input_abs}")"
  if [[ -d "${input_abs}" ]]; then
    docker_args+=(-v "${input_abs}:/input${index}:ro")
    container_inputs+=("/input${index}")
  else
    input_dir="$(dirname "${input_abs}")"
    docker_args+=(-v "${input_dir}:/input${index}:ro")
    container_inputs+=("/input${index}/${input_name}")
  fi
done

if [[ -n "${password}" ]]; then
  docker_args+=(-e "EXCEL2CSV_PASSWORD=${password}")
fi

docker_args+=("${IMAGE}")
docker_args+=("${container_inputs[@]}")
docker_args+=(-o "/output/${output_name}")

if [[ "${gzip_output}" -eq 1 ]]; then
  docker_args+=(--gzip)
fi

docker "${docker_args[@]}"
printf 'Wrote %s\n' "${output_abs}"

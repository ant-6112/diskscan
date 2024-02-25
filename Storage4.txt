#!/bin/bash

# Set default values
path="${1:-.}"
minimum_size="${2:-10.0}"
unit="${3:-MB}"

# Check if the unit entered is valid
valid_units=("MB" "KB" "GB")
if [[ ! " ${valid_units[@]} " =~ " $unit " ]]; then
  echo "Unit entered is invalid. Please enter a valid unit from ${valid_units[*]}"
  exit 1
fi

# Multiplier based on unit
case $unit in
  MB) multiplier=$((1024**2));;
  KB) multiplier=$((1024));;
  GB) multiplier=$((1024**3));;
esac

# Size threshold in kilobytes
size_threshold_kilobytes=$(awk -v size="$minimum_size" 'BEGIN {print size * 1024}')

# Find large files and print their information
find "$path" -type f -not -name "passwd" -exec du -k {} + | while read -r size file; do
  if (( size * 1024 > size_threshold_kilobytes )); then
    if [[ "$file" != "/etc/passwd" ]]; then
      user_id=$(stat -c %u "$file" 2>/dev/null)
      user_name=$(awk -v uid="$user_id" -F ":" '{if ($3 == uid) print $1}' /etc/passwd)
      echo "$file: $(awk -v size="$size" -v unit="$unit" -v multiplier="$multiplier" 'BEGIN {print size * 1024 / multiplier " " unit}') (Created by: ${user_name:-Unknown})"
    fi
  fi
done

# Print top users by storage usage
echo "Top Users by Storage Usage:"
find "$path" -type f -exec stat -c "%U %s" {} + | awk '{user_storage[$1]+=$2} END {for (user in user_storage) {printf "%s: %.2f %s\n", user, user_storage[user]/'$multiplier', "'$unit'"}}' | sort -k2,2nr

#!/bin/bash

if [[ "$#" -eq 1 ]]; then
  BASEURL="$1"
else
  echo "This script requires 1 argument, a base URL."
  exit 1
fi

# Purge existing rooms
curl $BASEURL/purge
echo

# Create set of new rooms
IFS="|"
while read room capacity schedule takers
do
    curl -H "Content-Type: application/json" --data-raw "{\"name\":\"$room\",\"capacity\":$capacity,\"schedule\":\"$schedule\",\"takers\":\"$takers\"}" $BASEURL/save
    echo
done <<EOF
01Mercury|4|0000000000000000|:::::::::::::::
02Venus|6|0000000000000000|:::::::::::::::
03Earth|8|0000000000000000|:::::::::::::::
04Mars|10|0000000000000000|:::::::::::::::
05Jupiter|12|0000000000000000|:::::::::::::::
06Saturn|14|0000000000000000|:::::::::::::::
07Uranus|16|0000000000000000|:::::::::::::::
08Neptune|18|0000000000000000|:::::::::::::::
09Pluto|3|0000000000000000|:::::::::::::::
EOF






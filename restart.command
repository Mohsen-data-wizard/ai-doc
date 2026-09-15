#!/bin/bash
cd "$(dirname "$0")"
echo "Restarting Agents Office..."
lsof -ti :4520 | xargs -r kill
sleep 1
npm start

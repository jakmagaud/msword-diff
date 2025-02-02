#!/bin/bash

if command -v python3 &>/dev/null; then
    echo "Python 3 is installed: $(python3 --version)"
else
    echo "Installing Python 3..."
    brew install python3
fi

python3 -m venv venv
source venv/bin/activate

# Install requirements if file exists
if [ -f "requirements.txt" ]; then
    pip install -r requirements.txt
    echo "Dependencies installed from requirements.txt"
else
    echo "No requirements.txt found"
fi

python3 script_grade.py

# Setup
For windows, use the following commands for setup:

>python -m venv .venv
>
>.venv\Scripts\activate
>
>pip install -r requirements.txt

*Can remove gunicorn from requirements.txt if only a local dev server is required.<br><br>

# Running a local server
After setup, run "flask --app app run" in the project directory to run a local server.

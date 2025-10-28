Joval Wines NIST Playbook Tracker - Local Docker (Full)

Instructions:
1. Place your .docx files in the 'playbooks' folder.
2. Build the Docker image:
   docker build -t jovalwines-nist-playbook:latest .
3. Run the container:
   docker run -p 8501:8501 jovalwines-nist-playbook:latest
4. Open http://localhost:8501

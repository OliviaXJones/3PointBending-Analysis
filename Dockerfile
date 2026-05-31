FROM python:3.11-slim

WORKDIR /app

# System dependencies for matplotlib headless rendering
RUN apt-get update && apt-get install -y --no-install-recommends \
    libgomp1 \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies first (layer caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy only the files the app needs
COPY streamlit_app.py .
COPY SingleStudy_workflow.py .
COPY .streamlit/ .streamlit/

EXPOSE 8080

CMD ["streamlit", "run", "streamlit_app.py", \
     "--server.port=8080", \
     "--server.address=0.0.0.0", \
     "--server.headless=true", \
     "--server.enableCORS=false"]

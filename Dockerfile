# nextcloud-mcp mit Office-Werkzeugen (bos_office.py): Dokumente, Tabellen, Praesentationen und PDF in der
# Nextcloud anlegen, bearbeiten und umwandeln. Basis fest auf nextcloud-mcp-server 0.212.0; sie bringt
# python-docx, openpyxl und python-pptx mit, es fehlt nur pypdf.
FROM ghcr.io/cbcoutinho/nextcloud-mcp-server@sha256:8fbf449a68e92b416e475804dd4adaf4f9c43a5b3d83e8dc0da0b112b359700e
USER root
RUN uv pip install --python /opt/venv/bin/python --no-cache "pypdf==6.21.4"
COPY bos_office.py bos_start.py /opt/bos/
ENV PYTHONPATH=/opt/bos
USER appuser
ENTRYPOINT ["/opt/venv/bin/python", "/opt/bos/bos_start.py", "run", "--host", "0.0.0.0"]
CMD ["--transport", "streamable-http", "--oauth", "--port", "8000"]

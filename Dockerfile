# nextcloud-mcp mit Office-Werkzeugen (bos_office.py): Dokumente, Tabellen, Praesentationen und PDF in der
# Nextcloud anlegen, bearbeiten und umwandeln. Basis fest auf nextcloud-mcp-server 0.202.1; sie bringt
# python-docx, openpyxl und python-pptx mit, es fehlt nur pypdf.
FROM ghcr.io/cbcoutinho/nextcloud-mcp-server@sha256:b842475624063f5c9e366a7af0e571848bd79bf2f6e345baac83ecf6bd9ae9e1
USER root
RUN uv pip install --python /opt/venv/bin/python --no-cache "pypdf==6.21.4"
COPY bos_office.py bos_start.py /opt/bos/
ENV PYTHONPATH=/opt/bos
USER appuser
ENTRYPOINT ["/opt/venv/bin/python", "/opt/bos/bos_start.py", "run", "--host", "0.0.0.0"]
CMD ["--transport", "streamable-http", "--oauth", "--port", "8000"]

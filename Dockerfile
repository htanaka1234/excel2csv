FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

COPY pyproject.toml README.md /app/
COPY src /app/src
COPY tests /app/tests

RUN pip install --upgrade pip \
    && pip install ".[test]"

ENTRYPOINT ["excel2csv"]

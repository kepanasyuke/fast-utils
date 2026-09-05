PORT ?= 8000
HOST ?= 0.0.0.0
PYTHON ?= python3

.PHONY: help install run clean kill-port

help:
	@echo "Доступные команды:"
	@echo "  make install  - Установить зависимости через pip"
	@echo "  make run      - Установить зависимости и запустить сервер"
	@echo "  make clean    - Удалить сгенерированные презентации"

install:
	$(PYTHON) -m pip install --upgrade pip
	$(PYTHON) -m pip install -r requirements.txt

run: install
	$(PYTHON) app.py

clean:
	rm -f generated_*.pptx generated_presentations/*.pptx

kill-port:
	@echo "Killing process on port 8000..."
	@lsof -ti :8000 | xargs kill -9 || echo "No process on port 8000"
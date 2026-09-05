PORT ?= 8000
HOST ?= 0.0.0.0
PYTHON ?= python3

.PHONY: help install run clean

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
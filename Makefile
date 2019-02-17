build: clean
	pyinstaller --onefile --windowed docx-merge.py

clean:
	rm -rf __pycache__ build dist *.spec

.PHONY: build
.PHONY: clean

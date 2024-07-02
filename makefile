run:
	python3 widget.py
ui:
	pyside6-uic form.ui -o ui_form.py
build:
	pyinstaller -F -w widget.py -n Idiom
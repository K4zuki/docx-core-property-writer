html:
	cd docs; \
	make initdir html

install:
	pip3 install --break-system-packages .

uninstall:
	pip3 uninstall -y docx-core-property-writer

reinstall: uninstall install

clean:
	cd docs; \
	make initdir clean

tex:
	cd docs; \
	make initdir tex

docx:
	cd docs; \
	make initdir docx

reverse-docx:
	cd docs; \
	make initdir reverse-docx

pdf:
	cd docs; \
	make initdir pdf

wheel:
	sudo python3 setup.py bdist_wheel

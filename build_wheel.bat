@echo off
pipenv install wheel
pipenv run python.exe setup.py bdist_wheel
pause
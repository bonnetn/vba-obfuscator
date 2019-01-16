# VBA Obfuscator

This is a school project. We wanted to implement an obfuscator for VBA scripts. 

## Run it with docker

 If you have docker installed you can directly do:

 > cat YOUR_MACRO.vbs | docker run -i --rm bonnetn/vba-obfuscator /dev/stdin

## Run it in a python environment

If you do not have docker, you can install the dependencies yourself:

To launch this project you must install *pygments*.
> pip install -r requirements.txt

Then run:
> python3 obfuscate.py YOUR_MACRO.vbs

## Demo

[![asciicast](https://asciinema.org/a/5Ptyf5oNGT7xtkZZvnqNDHMml.svg)](https://asciinema.org/a/5Ptyf5oNGT7xtkZZvnqNDHMml)

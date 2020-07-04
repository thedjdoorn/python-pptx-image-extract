# python-pptx-image-extract
This is a tool for extracting images from PowerPoint Presentations. It is very useful for batch processing.

## Installation
This project depends on *opencv*, *pyton-pptx* and *numpy*. The latter two can be installed with pip, for opencv it is recommended to install using your system package manager (you're looking for the `python3-opencv` package).

## Usage
Please note this only works with *pptx* presentations due to limitations in the *python-pptx* library. Presentations can be easily converted using the *libreoffice-impress* CLI.

Images can be extracted by running `python3 extract.py presentation.pptx`, which puts them in a directory called *presentation*, a custom directory can be specified using **--name directory**. The images can also be inverted, by passing **-i** or **--invert**.
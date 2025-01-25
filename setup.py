from setuptools import setup, find_packages

setup(
    name='iGOT Reports',
    version='0.0.1',
    author='JaanuShaikDev',
    author_email='johnsyda.shaik@gmail.com',
    install_requires=[
        'pandas',
        'openpyxl',
        'plotly',
        'importlib-metadata;python_version < "3.12"'    
    ],
    packages = find_packages()
)
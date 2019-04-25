from setuptools import setup

setup(
    name = "MicroBill",
    version = "1.0.0",
    author = "Juan Barbosa",
    author_email = "js.barbosa10@uniandes.edu.co",
    # description = ("Abacus Software is a suite of tools build to ensure your experience with Tausand's coincidence counters becomes simplified."),
    license = "GPL",
    keywords = "example documentation tutorial",
    url = "https://github.com/jsbarbosa/microbill",
    packages=['microbill'],
    install_requires = ['numpy', 'PyQt5', 'pandas', 'xlrd', 'openpyxl',
        'xlsxwriter', 'reportlab', 'html2text', 'lxml'
    ],
    long_description = "",#read('README'),
    classifiers = [
        "Development Status :: 1 - Planning",
        "Topic :: Utilities",
        "License :: OSI Approved :: GNU General Public License (GPL)",
    ],
    data_files=[('microbill/Registers', ['microbill/Registers/Clientes.xlsx',
                            'microbill/Registers/Independientes.xlsx',
                            'microbill/Registers/Precios.xlsx',
                            'microbill/Registers/Registro.xlsx'])
                            ]
)

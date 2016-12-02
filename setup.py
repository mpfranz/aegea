"""
Using Cython
1. Change the file name in line 5 to the module to compile
2. Save the module to compile as .pyx
3. Bash: python setup.py build_ext --inplace
4. A .so file should now exist in the directory
5. To use the compiled module use: import <module name>
6. Use the module as you would use any: <module name>.<method name>
"""

from distutils.core import setup
from Cython.Build import cythonize

setup(
    ext_modules = cythonize("optionpricing_tte.pyx")
)
import sys
import runpy

if __name__ == '__main__':
    sys.argv[0] = 'esptool'
    runpy.run_module('esptool', run_name='__main__', alter_sys=True)

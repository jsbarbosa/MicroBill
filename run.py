import importlib
import microbill

if __name__ == '__main__':
    exit_code = microbill.run()
    while exit_code == microbill.constants.EXIT_CODE_REBOOT:
        importlib.reload(microbill)
        exit_code = microbill.run()

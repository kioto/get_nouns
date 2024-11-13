"""
"""
import sys


def main(filename):
    print(filename)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print(f'Usage: python {sys.argv[0]} <text-file> | <excel-file>')
        exit()


    main(sys.argv[1])

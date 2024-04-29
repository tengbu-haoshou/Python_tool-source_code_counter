#!/usr/bin/env python3

"""
Hello world
"""

# import
import sys

# main
"""main"""
def main():

    # hello world
    str='\'hello world\''       # string
    str="'hello world'"         # string

    # print
    print(str)                  # print

    # end
    sys.exit(0)                 # end

# goto main
if __name__=='__main__':
    main()

# irregular (Not processed properly)
    'string
    '
    "string
    """comment"
    """comment""
    "
    ""
    """

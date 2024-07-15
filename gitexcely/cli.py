import traceback
import argparse
import sys
from .main import apply, convert, init, clone

def cli():
    parser = argparse.ArgumentParser()
    parser.add_argument('action', type=str, help='action for the cli tool', choices=['init', 'clone', 'apply', 'to_excel', 'to_text', 'auto_convert'])
    parser.add_argument('--path', type=str, help='path to make action', default='.')
    parser.add_argument('--url', type=str, help='url to clone project and apply git excel', default='.')
    args = parser.parse_args()
    try:
        if args.action=='init':
            init(args.path)
        elif args.action=='clone':
            clone(args.path, args.url)
        elif args.action=='apply':
            apply(args.path)
        elif args.action=='auto_convert' or args.action=='to_text' or args.action=='to_excel':
            convert(args.action, args.path)
    except Exception as error:
        traceback.print_exc()
        print('------')
        print(error)
        print("Error!")
        sys.exit(1)
    else:
        print("Succeed!")
        sys.exit(0)

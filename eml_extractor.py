import re
from argparse import ArgumentParser, ArgumentTypeError
from email import message_from_file, policy
from pathlib import Path
from typing import List


def extract_attachments(file: Path, destination: Path) -> None:
    print(f'PROCESSING FILE "{file}"')
    with file.open() as f:
        email_message = message_from_file(f, policy=policy.default)
        email_subject = email_message.get('Subject')
        
        # Sanitize the folder name, providing a default if it's None
        folder_name = sanitize_foldername(email_subject)
        if folder_name is None:
            print(">> Skipping file due to missing or invalid folder name.")
            return
        
        basepath = destination / folder_name
        
        # Rest of your code...
        # ignore inline attachments
        attachments = [item for item in email_message.iter_attachments() if item.is_attachment()]  # type: ignore
        if not attachments:
            print('>> No attachments found.')
            return
        for attachment in attachments:
            filename = attachment.get_filename()
            if filename is None:                
                filename = attachment.get_payload(i = 0).get_payload(i = 1).get_filename()
                print(f'>> Attachment found: {filename} \n')                
                filepath = basepath / filename
                payload = attachment.get_payload(i = 0).get_payload(i = 1).get_payload(decode=True)
                if filepath.exists():
                    overwrite = input(f'>> The file "{filename}" already exists! Overwrite it (Y/n)? ')
                    save_attachment(filepath, payload) if overwrite.upper() == 'Y' else print('>> Skipping...')
                else:
                    basepath.mkdir(exist_ok=True)
                    save_attachment(filepath, payload)
                    
            elif('eml' in filename):
                print(f'>> Attachment found: {filename}')
                filepath = basepath / filename
                payload = attachment.get_payload(decode=True)
                if filepath.exists():
                    overwrite = input(f'>> The file "{filename}" already exists! Overwrite it (Y/n)? ')
                    save_attachment(filepath, payload) if overwrite.upper() == 'Y' else print('>> Skipping...')
                else:
                    basepath.mkdir(exist_ok=True)
                    save_attachment(filepath, payload)
                extract_attachments(filepath, basepath)
            
            else:
                print(f'>> Attachment found: {filename}')
                filepath = basepath / filename
                payload = attachment.get_payload(decode=True)
                if filepath.exists():
                    overwrite = input(f'>> The file "{filename}" already exists! Overwrite it (Y/n)? ')
                    save_attachment(filepath, payload) if overwrite.upper() == 'Y' else print('>> Skipping...')
                else:
                    basepath.mkdir(exist_ok=True)
                    save_attachment(filepath, payload)








def sanitize_foldername(name: str) -> str:
    illegal_chars = r'[/\\|\[\]\{\}:<>+=;,?!*"~#$%&@\']'
    return re.sub(illegal_chars, '_', name)

def save_attachment(file: Path, payload: bytes) -> None:
    with file.open('wb') as f:
        print(f'>> Saving attachment to "{file}"')
        f.write(payload)

def get_eml_files_from(path: Path, recursively: bool = False) -> List[Path]:
    if recursively:
        return list(path.rglob('*.eml'))
    return list(path.glob('*.eml'))

def check_file(arg_value: str) -> Path:
    file = Path(arg_value)
    if file.is_file() and file.suffix == '.eml':
        return file
    raise ArgumentTypeError(f'"{file}" is not a valid EML file.')

def check_path(arg_value: str) -> Path:
    path = Path(arg_value)
    if path.is_dir():
        return path
    raise ArgumentTypeError(f'"{path}" is not a valid directory.')

def get_argument_parser():
    parser = ArgumentParser(
        usage='%(prog)s [OPTIONS]',
        description='Extracts attachments from .eml files'
    )
    # force the use of --source or --files, not both
    source_group = parser.add_mutually_exclusive_group()
    source_group.add_argument(
        '-s',
        '--source',
        type=check_path,
        default=Path.cwd(),
        metavar='PATH',
        help='the directory containing the .eml files to extract attachments (default: current working directory)'
    )
    parser.add_argument(
        '-r',
        '--recursive',
        action='store_true',
        help='allow recursive search for .eml files under SOURCE directory'
    )
    source_group.add_argument(
        '-f',
        '--files',
        nargs='+',
        type=check_file,
        metavar='FILE',
        help='specify a .eml file or a list of .eml files to extract attachments'
    )
    parser.add_argument(
        '-d',
        '--destination',
        type=check_path,
        default=Path.cwd(),
        metavar='PATH',
        help='the directory to extract attachments to (default: current working directory)'
    )
    return parser

def parse_arguments():
    parser = get_argument_parser()
    return parser.parse_args()

def main():
    args = parse_arguments()

    eml_files = args.files or get_eml_files_from(args.source, args.recursive)
    if not eml_files:
        print(f'No EML files found!')

    for file in eml_files:
        extract_attachments(file, destination=args.destination)
    print('Done.')


if __name__ == '__main__':
    main()

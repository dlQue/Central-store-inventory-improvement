import os
import hashlib

def get_file_hash(file_path):
    """Calculate SHA-256 hash of a file."""
    sha256_hash = hashlib.sha256()
    with open(file_path, 'rb') as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()

def delete_duplicate_pdfs(folder_path):
    """Delete duplicate PDF files in the given folder."""
    hashes = {}
    files_to_delete = []
    count = 0  # Initialize count within the function

    print("Checking for duplicates...")

    for filename in os.listdir(folder_path):
        if filename.endswith(('.pdf', '.PDF')):  # Correct syntax for multiple file extensions
            file_path = os.path.join(folder_path, filename)
            file_hash = get_file_hash(file_path)

            if file_hash in hashes:
                files_to_delete.append(file_path)
            else:
                hashes[file_hash] = file_path

    for file_path in files_to_delete:
        os.remove(file_path)
        count += 1
        print(f"Deleted duplicate file: {file_path}")

    return count  # Return the count of deleted files

# Replace 'your/folder/path' with the actual path to your folder containing the PDF files
deleted_count = delete_duplicate_pdfs('/Users/quedonglin/Downloads/Aurora invoices (Special Order) 2')

print('Done.')
print('Deleted file count:', deleted_count)

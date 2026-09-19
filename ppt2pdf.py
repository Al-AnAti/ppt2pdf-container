import subprocess
import argparse
from pathlib import Path


def convert_file(file_path):
    outdir = file_path.parent
    try:
        print(f"converting: {file_path.name}...")
        subprocess.run(
            ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', str(outdir), str(file_path)],
            check=True
        )
        print(f"success: {file_path.name}.pdf")
    except subprocess.CalledProcessError as e:
        print(f"Error converting {file_path.name}: {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("file", help="The PPT file to convert")
    target_path = Path(parser.parse_args().file)

    if not target_path.exists():
        print(f"Error: The path '{target_path}' does not exist.")
    elif target_path.is_file() and target_path.suffix.lower() in ['.ppt', '.pptx']:
        convert_file(target_path)
    elif target_path.is_dir():
        for file in target_path.iterdir():
            if file.is_file() and file.suffix.lower() in ['.ppt', '.pptx']:
                convert_file(file)
    else:
        print("Error: Please provide a valid .ppt/.pptx file or a directory containing them.")
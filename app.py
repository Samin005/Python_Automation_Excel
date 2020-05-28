from pathlib import Path

path = Path()
# to view existing files in current directory
for file in path.glob('*'):
    print(file)

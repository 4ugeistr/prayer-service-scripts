import easygui
import os

def find_docx_files(folder_path):
  """Finds all *.docx files in the specified folder and its subfolders.

  Args:
    folder_path: The path to the folder where the search should begin.

  Returns:
    A list of full paths to all the found *.docx files.
  """

  docx_files = []
  for root, _, files in os.walk(folder_path):
    for filename in files:
      if filename.endswith(".docx"):
        full_path = os.path.join(root, filename)
        docx_files.append(full_path)
  return docx_files

def main():
  """Prompts the user to choose a folder and displays the found *.docx files."""

  # Use easygui for a simple dialog
  folder_path = easygui.diropenbox(title="Select Folder")  # Open folder selection dialog

  if folder_path:  # Check if a folder was selected
    docx_files = find_docx_files(folder_path)
    if docx_files:
      print(docx_files)

if __name__ == "__main__":
  main()

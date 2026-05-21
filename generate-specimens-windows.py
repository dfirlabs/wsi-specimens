#!/usr/bin/env python3
# pylint: disable=invalid-name
"""Script to generate Windows shell item test files.

Requires Windows and pywin32.
"""

import os
import shutil

import pywintypes

from win32com.shell import shell
from win32com.shell import shellcon


class WindowsShellItemGenerator:
  """Windows shell item generator."""

  # pylint: disable=redefined-outer-name

  def __init__(self, specimens_path):
    """Initializes a Windows shell item generator.

    Args:
      specimens_path (str): path of the directory where to store the generated specimens.
    """
    super().__init__()
    self._desktop_shell_folder = shell.SHGetDesktopFolder()
    self._specimens_path = specimens_path

  def GetItemListByPath(self, path):
    """Retrieves an item list for a specific path.

    Args:
      path (str): path.

    Returns:
      list[bytes]: shell item list or None not available.
    """
    shell_item_list, _ = shell.SHParseDisplayName(path, 0, None)
    return shell_item_list

  def GetSubItemsByPath(self, path):
    """Retrieves the sub items of a shell folder.

    Args:
      path (str): path.

    Returns:
      list[list[bytes]]: shell item lists of the sub items.
    """
    shell_item_list, _ = shell.SHParseDisplayName(path, 0, None)
    if not shell_item_list:
      return []

    try:
      shell_folder =  self._desktop_shell_folder.BindToObject(
          shell_item_list, None, shell.IID_IShellFolder)
    except pywintypes.com_error:
      return []

    sub_items = []

    try:
      enumerator = shell_folder.EnumObjects(
          shellcon.SHCONTF_FOLDERS |
          shellcon.SHCONTF_NONFOLDERS |
          shellcon.SHCONTF_INCLUDEHIDDEN)

      for sub_item in enumerator or []:
        sub_item_list = list(shell_item_list)
        sub_item_list.extend(sub_item)

        sub_items.append(sub_item_list)

    except pywintypes.com_error:
      return []

    return sub_items

  def WriteItemListSpecimen(self, name, shell_item_list):
    """Writes a shell item list specimen.

    Args:
      name (str): name of the specimen.
      shell_item_list (list[bytes]): shell item list.
    """
    path = os.path.join(self._specimens_path, f'{name:s}.bin')
    with open(path, 'wb') as file_object:
      for shell_item in shell_item_list:
        size = len(shell_item)
        file_object.write(size.to_bytes(2, 'little'))
        file_object.write(shell_item)

      file_object.write(int(0).to_bytes(2, 'little'))


if __name__ == '__main__':
  specimens_path = os.path.join(os.getcwd(), 'specimens')
  os.makedirs(specimens_path)

  generator = WindowsShellItemGenerator(specimens_path)

  # Generate a file entry (directory) shell item list.
  testdir_path = os.path.join(os.getcwd(), 'testdir')
  os.makedirs(testdir_path)

  shell_item_list = generator.GetItemListByPath(testdir_path)
  generator.WriteItemListSpecimen('file_entry1', shell_item_list)

  # Generate a file entry (file) shell item list.
  path = os.path.join(testdir_path, 'testfile.txt')
  with open(path, 'w', encoding='utf8') as file_object:
    file_object.write('Some text')

  shell_item_list = generator.GetItemListByPath(path)
  generator.WriteItemListSpecimen('file_entry2', shell_item_list)

  # Generate a compressed folder shell item list.
  path = os.path.join(os.getcwd(), 'test')
  shutil.make_archive(path, 'zip', 'testdir')

  path = os.path.join(os.getcwd(), 'test.zip')
  for index, sub_item_list in enumerate(
      generator.GetSubItemsByPath(path)):
    display_index = index + 1
    generator.WriteItemListSpecimen(
        f'compressed_folder{display_index:d}', sub_item_list)

  # Generate a control panel category shell item list.
  path = '::{26EE0668-A00A-44D7-9371-BEB064C98683}\\1'

  shell_item_list = generator.GetItemListByPath(path)
  generator.WriteItemListSpecimen('control_panel_category1', shell_item_list)

  # Generate control panel item shell item lists.
  for index, sub_item_list in enumerate(
      generator.GetSubItemsByPath(path)):
    display_index = index + 1
    generator.WriteItemListSpecimen(
        f'control_panel_item{display_index:d}', sub_item_list)

  # TODO: Generate a .search-ms file

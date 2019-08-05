import ast
import os
import time

from dstream_excel.tracker.timing import TimeTracker


class FileProcessTracker:

    def __init__(self, folder=None, restart=False, file_types=('csv',)):

        if folder is None:
            self.folder = os.getcwd()
        else:
            self.folder = os.path.abspath(folder)

        self.completed_list_path = os.path.join(self.folder, 'completed.txt')

        if restart:
            self.delete_completed_files()
        self.restart = restart

        self.load_completed_files()
        self.load_process_files(file_types=file_types)

    def file_generator(self):
        timer = TimeTracker(self.folder, restart=self.restart)
        num_items = len(self.process_list)

        for file in self.process_list:
            yield os.path.join(self.folder, file)
            self.add_file_to_completed(file)
            timer.time_estimate(num_items)

    def add_file_to_completed(self, file):
        self.completed_list.extend([file])
        _update_completed_files(self.completed_list_path, self.completed_list)

    def load_completed_files(self):
        self.completed_list = _load_completed_files(self.completed_list_path)

    def load_process_files(self, file_types):
        self.process_list = _load_to_process_files(self.folder, self.completed_list, file_types)

    def delete_completed_files(self):
        _delete_completed_files(self.completed_list_path)


def _load_to_process_files(folder, completed_list, file_types):
    files = _load_initial_file_list(folder, file_types)
    return [file for file in files if file not in completed_list]


def _update_completed_files(completed_list_path, completed_list):
    _write_to_file_with_retries(completed_list_path, completed_list)


def _write_to_file_with_retries(*args, retries_remaining=10, **kwargs):
    try:
        with open(args[0], 'w') as f:
            f.write(f'{args[1]}')
    except (OSError, PermissionError):
        time.sleep(.1)
        _write_to_file_with_retries(*args, retries_remaining=retries_remaining-1, **kwargs)


def _load_completed_files(completed_list_path):

    # Not started yet, none completed
    if not os.path.exists(completed_list_path):
        return []

    with open(completed_list_path, 'r') as f:
        completed_list = ast.literal_eval(f.read())

    return completed_list


def _load_initial_file_list(folder, file_types):
    return [file for file in next(os.walk(folder))[2] if any([file.endswith(ending) for ending in file_types])]


def _delete_completed_files(completed_list_path):
    if os.path.exists(completed_list_path):
        os.remove(completed_list_path)
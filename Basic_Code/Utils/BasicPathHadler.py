import os


class Singleton(type):
    _instances = {}

    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super(Singleton, cls).__call__(*args, **kwargs)
        return cls._instances[cls]


class BasicPathHandler(metaclass=Singleton):
    def __init__(self):
        self._project_name_lowercase = "slothandler"
        self._games_result_data_folder = "Game_Result_Data"
        self._games_source_data_folder = "Game_Source_Data"
        self._project_folder = ""

    def _FindProjectFolderPath(self):
        dir_path_prev, base_name_prev = os.path.split(os.getcwd())
        while True:
            dir_path_new, base_name_new = os.path.split(dir_path_prev)
            if base_name_new.lower() == self._project_name_lowercase:
                self._project_folder = dir_path_prev
                break
            else:
                dir_path_prev, base_name_prev = dir_path_new, base_name_new

    def _ProcessRelativeFolderPath(self, relative_folder_path: str, absolute_head_path: str):
        path, base = os.path.split(relative_folder_path)
        folders = []
        while base != "":
            folders.append(base)
            path, base = os.path.split(path)
        folders.reverse()

        for folder in folders:
            absolute_head_path = os.path.join(absolute_head_path, folder)
            if not os.path.exists(absolute_head_path):
                os.makedirs(absolute_head_path)
        return absolute_head_path

    def GetPARPath(self, game_name: str, file_name: str):
        self._FindProjectFolderPath()
        relative_folder_path = os.path.join(self._games_result_data_folder, game_name)
        absolute_save_file_path = self._project_folder

        return os.path.join(self._ProcessRelativeFolderPath(relative_folder_path, absolute_save_file_path), file_name)

    def GetResultDataFilePath(self, game_name: str, inner_folder: str, file_name: str):
        self._FindProjectFolderPath()
        relative_folder_path = os.path.join(self._games_result_data_folder, game_name, inner_folder)
        absolute_save_file_path = self._project_folder

        return os.path.join(self._ProcessRelativeFolderPath(relative_folder_path, absolute_save_file_path), game_name)

    def GetTempGraphsFolderPath(self, game_name_short: str, version_name: str, folder_name="tmp_par_graphs"):
        self._FindProjectFolderPath()
        relative_folder_path = os.path.join(self._games_result_data_folder, game_name_short, version_name, folder_name)
        absolute_save_file_path = self._project_folder

        return self._ProcessRelativeFolderPath(relative_folder_path, absolute_save_file_path)

    def GetJsonStatsPath(self, game_name_short: str, version_name: str, tags: list = ['usual']):
        self._FindProjectFolderPath()
        json_stats_folder_path = os.path.join(self._project_folder, self._games_source_data_folder, game_name_short, version_name)

        res = ''
        current_directory = os.getcwd()
        os.chdir(json_stats_folder_path)
        for file_name in os.listdir():
            if all([tag in file_name for tag in tags]) and '.json' in file_name:
                res = os.path.join(json_stats_folder_path, file_name)
        os.chdir(current_directory)

        if res == '':
            raise NameError('NO NEEDED JSON STATS IN FOLDER')

        return res



if __name__ == "__main__":
    t = BasicPathHandler()
    r = BasicPathHandler()

    print(id(t))
    print(id(r))
    print(t.GetPARPath("MIN10", "PAR.xlsx"))
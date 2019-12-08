from modules.packages.dll import Dll


class Dll64(Dll):
    """DLL x64 analysis package."""

    def start(self, path):
        return self._execute_exports(path, '64')

class MSDocument:
    def __init__(self, path: str):
        with open(path, "r") as f:
            self.code = f.read()
            self.doc_var = {}

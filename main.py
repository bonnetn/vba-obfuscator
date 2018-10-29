from obfuscator.modifier.base import Pipe
from obfuscator.modifier.change_variable_names import ChangeVariableNames
from obfuscator.modifier.crypt_strings import CryptStrings
from obfuscator.msdocument import MSDocument

VBA_PATH = "example_macro/download_payload.vba"

if __name__ == "__main__":
    doc = MSDocument(VBA_PATH)
    Pipe(doc).run(
        ChangeVariableNames(),
        CryptStrings()
    )
    print(doc.code)

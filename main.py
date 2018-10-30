from obfuscator.modifier.base import Pipe
from obfuscator.modifier.randomize_function_names import RandomizeFunctionNames
from obfuscator.modifier.randomize_variable_names import RandomizeVariableNames
from obfuscator.modifier.crypt_strings import CryptStrings
from obfuscator.msdocument import MSDocument

VBA_PATH = "example_macro/download_payload.vba"

if __name__ == "__main__":
    doc = MSDocument(VBA_PATH)
    Pipe(doc).run(
        RandomizeVariableNames(),
        CryptStrings(),
        RandomizeFunctionNames(),
    )
    print(doc.code)

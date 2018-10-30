import logging
import re
from typing import Iterable

from obfuscator.modifier.base import Modifier
from obfuscator.msdocument import MSDocument
from obfuscator.util import get_random_string_of_random_length, replace_whole_word

LOG = logging.getLogger(__name__)
FIND_FUNC_NAMES_REGEX = r"\s*(?:Sub|Function)\s+\b(\w+)\s*\("

FUNCTION_BLACKLIST = {
    "Auto_Open",
    "AutoOpen",
    "Workbook_Open",
}


class RandomizeFunctionNames(Modifier):
    def run(self, doc: MSDocument) -> None:
        for name in _get_function_names(doc.code):
            if name not in FUNCTION_BLACKLIST:
                new_name = get_random_string_of_random_length()
                doc.code = replace_whole_word(doc.code, name, new_name)
                LOG.debug("Replaced '{}' with '{}'.".format(name, new_name))


def _get_function_names(content: str) -> Iterable[str]:
    func_names = re.finditer(FIND_FUNC_NAMES_REGEX, content)
    return map(lambda v: v.group(1), func_names)

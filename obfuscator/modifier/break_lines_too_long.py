import logging

from pygments import highlight
from pygments.formatter import Formatter
from pygments.lexers.dotnet import VbNetLexer
from pygments.token import Token

from obfuscator.modifier.base import Modifier
from obfuscator.msdocument import MSDocument

MAX_LINE_WIDTH = 1000

LOG = logging.getLogger(__name__)


def _do_split_line(line: str) -> str:
    return highlight(line, VbNetLexer(), _BreakLinesTooLong())


def _split_line_if_necessary(line: str) -> str:
    if len(line) >= MAX_LINE_WIDTH:
        LOG.info("Line '{:.30s}[...]' is too long.".format(line))
        return _do_split_line(line)
    return line


class BreakLinesTooLong(Modifier):
    def run(self, doc: MSDocument) -> None:
        code = doc.code
        prev = ""
        while prev != code:
            prev = code
            lines_code = code.split("\n")
            lines_code = map(_split_line_if_necessary, lines_code)
            code = "\n".join(lines_code)

        doc.code = code


class _BreakLinesTooLong(Formatter):
    def format(self, tokensource, outfile):
        break_points = []
        last_type = None
        last_val = ''
        line = ''
        for ttype, value in tokensource:
            if last_type != ttype or (ttype == Token.Punctuation and value in ",+&"):
                if last_type in {Token.Punctuation}:
                    break_points.append(len(line) - len(last_val))

                last_val = ''
                last_type = ttype

            last_val += value
            line += value

        for i in range(len(break_points) - 1):
            bp1 = break_points[i]
            bp2 = break_points[i + 1]
            if bp1 < MAX_LINE_WIDTH <= bp2:
                outfile.write(line[:bp1+1] + " _\n" + line[bp1+1:])
                break

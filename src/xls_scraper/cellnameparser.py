import abc
import re

__all__ = ["CellRangeParser", "CellNameParser"]


def _parse_cell_name(cel_name: str) -> tuple[str, int]:
    c = cel_name.rstrip('0123456789')
    r = cel_name[len(c):]
    return _col_title_to_number(c), int(r)


def _col_title_to_number(col_title: str) -> int:
    ans = 0
    for i in col_title:
        ans = ans * 26 + (ord(i) - ord('A')) + 1
    return ans


class CellParserStrategy(metaclass=abc.ABCMeta):

    @abc.abstractmethod
    def valid(self, cell_name: str) -> bool:
        pass

    @abc.abstractmethod
    def parse(self, cell_name: str):
        pass


class CellRangeParser(CellParserStrategy):
    __RANGE_PATTERN__ = r'^[A-Z]{1,3}[1-9]{1}[0-9]{0,5}:[A-Z]{1,3}[1-9]{1}[0-9]{0,5}$'

    def valid(self, cell_name: str) -> bool:
        return re.match(self.__RANGE_PATTERN__, cell_name)

    def parse(self, cell_name: str):
        fr, to = tuple(cell_name.split(':'))
        from_tuple = _parse_cell_name(fr)
        to_tuple = _parse_cell_name(to)
        return dict(range_beg=from_tuple, range_end=to_tuple)


class CellNameParser(CellParserStrategy):
    __CELLNAME_PATTERN__ = r'^[A-Z]{1,3}[1-9]{1}[0-9]{0,5}$'

    def valid(self, cell_name: str) -> bool:
        return re.match(self.__CELLNAME_PATTERN__, cell_name)

    def parse(self, cell_name: str):
        return dict(col_row=_parse_cell_name(cell_name))


class CellSelectorParser:
    __parser_strgs: list[CellParserStrategy] = [
        CellNameParser(), CellRangeParser()]

    def get_parser(self, cell_name)->CellParserStrategy:
        valid_strg = next(
            (strg for strg in self.__parser_strgs if strg.valid(cell_name)),
            None)
        return valid_strg

    def parse_selection(self, selection:str)->dict:
        if strg:=self.get_parser(selection):
            return strg.parse(selection)
        else:
            return None

    def parse_selection(self, strategy: CellParserStrategy, selection:str)->dict:
        if strategy:
            return strategy.parse(selection)
        else:
            return None
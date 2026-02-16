"""Excel translation package."""

__all__ = ["ProcessingResult", "process_excel_file"]


def __getattr__(name: str):
    if name in {"ProcessingResult", "process_excel_file"}:
        from .processor import ProcessingResult, process_excel_file

        return {"ProcessingResult": ProcessingResult, "process_excel_file": process_excel_file}[name]
    raise AttributeError(name)

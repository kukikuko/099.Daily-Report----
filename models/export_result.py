from dataclasses import dataclass


@dataclass(slots=True)
class ExportResult:
    pdf_success: bool = False
    png_success: bool = False
    pdf_path: str = ""
    png_path: str = ""
    error_message: str = ""

    @property
    def is_full_success(self) -> bool:
        return self.pdf_success and self.png_success

    @property
    def is_partial_success(self) -> bool:
        return self.pdf_success and not self.png_success

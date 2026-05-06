from pathlib import Path


class EmailService:
    """
    Stub/base para a etapa de e-mail.

    Quando tiver o acesso real, implemente aqui:
    - IMAP
    - Outlook local
    - Microsoft Graph
    - Gmail API

    A saída esperada desse serviço deve ser uma lista de arquivos baixados.
    """

    def baixar_anexos(self) -> list[Path]:
        raise NotImplementedError(
            "Etapa de e-mail ainda não implementada. "
            "Por enquanto, rode informando --arquivo com Excel/CSV local."
        )
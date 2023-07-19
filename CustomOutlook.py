from RPA.Outlook.Application import Application
import gc


class CustomOutlook(Application):
    def __init__(self, autoexit: bool = True) -> None:
        super().__init__(autoexit)

    def quit_application(self, save_changes: bool = False) -> None:
        """Quit the application.

        :param save_changes: if changes should be saved on quit, default False
        """
        if not self.app:
            return
        self.close_document(save_changes)
        self.app.Quit()
        self.app = None
        gc.collect()

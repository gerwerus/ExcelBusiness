from PySide6.QtCore import Slot, Signal


class ProgressChangeMixin:
    changeProgressSignal = Signal(float)

    @Slot(float)
    def changeProgressSlot(self, percentage):
        self.progressBar.setValue(percentage)

class ConversoresExcel:
    def __init__(self):
        pass

    def conversorSegundosWialon(self, duration_str):
        parts = duration_str.split(':')
        if len(parts) == 3:
            self.partes = int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
            return self.partes
        elif len(parts) == 2:
            self.partes = int(parts[0]) * 60 + int(parts[1])
            return self.partes
        else:
            self.partes = int(parts[0])
            return self.partes


    def conversorSegundosUbicar(self, duration):
        parts = duration.split(' ')
        total_seconds = 0
        for part in parts:
            if 'h' in part:
                total_seconds += int(part.replace('h', '')) * 3600
            elif 'min' in part:
                total_seconds += int(part.replace('min', '')) * 60
            elif 's' in part:
                total_seconds += int(part.replace('s', ''))
        return self.total_seconds


    def conversorSegundosMDVR(self, duration_str):
        parts = duration_str.split(' ')
        minutes = 0
        seconds = 0
        for part in parts:
            if 'min' in part:
                minutes += int(part.replace('min', ''))
            if 's' in part:
                seconds += int(part.replace('s', ''))

        self.duracion = minutes * 60 + seconds
        return self.duracion
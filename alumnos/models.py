# alumnos/models.py
from django.db import models

class Curso(models.Model):
    nombre = models.CharField(max_length=150, unique=True)

    def __str__(self):
        return self.nombre

class Alumno(models.Model):
    nombre_completo = models.CharField(max_length=200)
    curso = models.ForeignKey(Curso, on_delete=models.CASCADE)

    def __str__(self):
        return f"{self.nombre_completo} ({self.curso})"


# modelos.py
class AsistenciaMensual(models.Model):
    alumno = models.ForeignKey(Alumno, on_delete=models.CASCADE)
    curso = models.ForeignKey(Curso, on_delete=models.CASCADE)
    mes = models.DateField()  # Usamos el día 1 del mes para referencia
    presentes = models.PositiveIntegerField(default=0)
    inasistentes = models.PositiveIntegerField(default=0)

    class Meta:
        unique_together = ('alumno', 'mes')

class DiasClaseMensual(models.Model):
    curso = models.ForeignKey(Curso, on_delete=models.CASCADE)
    mes = models.DateField()
    dias_clases = models.PositiveIntegerField(default=0)

    class Meta:
        unique_together = ('curso', 'mes')

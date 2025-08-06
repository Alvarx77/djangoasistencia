# alumnos/urls.py
from django.urls import path
from . import views
from .views import asistencia_mensual, ajax_actualizar_asistencia, ajax_actualizar_dias_clases, estadisticas, ajax_estadisticas_mes, reporte_cursos_mes

urlpatterns = [
    path('cargar_excel/', views.cargar_excel, name='cargar_excel'),
    path('lista_alumnos/', views.lista_alumnos, name='lista_alumnos'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('asistencia_mensual/', views.asistencia_mensual, name='asistencia_mensual'),
    path('ajax/actualizar_asistencia/', ajax_actualizar_asistencia, name='ajax_actualizar_asistencia'),
    path('ajax/actualizar_dias_clases/', ajax_actualizar_dias_clases, name='ajax_actualizar_dias_clases'),
    path('estadisticas/', estadisticas, name='estadisticas'),
    path('ajax/estadisticas_mes/', ajax_estadisticas_mes, name='ajax_estadisticas_mes'),
    path('reporte_cursos/', reporte_cursos_mes, name='reporte_cursos_mes'),
    path('exportar_excel/', views.exportar_excel, name='exportar_excel'),
    


]

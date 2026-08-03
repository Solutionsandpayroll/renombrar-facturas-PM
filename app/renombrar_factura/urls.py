from django.urls import path
from . import views

urlpatterns = [
    path("", views.inicio, name="inicio"),
    path("extraer-informacion/", views.cargar_excel, name="cargar_excel"),
    path("generar-reporte/", views.generar_reporte_final, name="generar_reporte_final"),
]
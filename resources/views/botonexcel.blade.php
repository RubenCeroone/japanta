<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Botón para Exportar a Excel</title>
    <!--Agrega los estilos de Bootstrap-->
    <link href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">
    <style>
        /* Estilos CSS personalizados */
        .btn-custom {
            /* Estilos personalizados para el botón */
            /* Por ejemplo, aquí definimos un fondo verde y letras blancas */
            background-color: green;
            color: white;
            height: 100%;
            margin: 0;
            display: flex;
            justify-content: center;
            align-items: center;
            top: 55%
            /* Otros estilos que desees aplicar */
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="row justify-content-center">
            <div class="col-md-6">
                <a href="{{ route('exportar_excel') }}" class="btn btn-success btn-custom">Exportar a Excel</a>
            </div>
        </div>
    </div>
</body>
</html>
Ejecutar desde local -> Doble consola

1. npm start
2. npm run electron

Lanzar un despliegue .exe

1. Actualizar la versión:
   - npm version <nueva-version> --no-git-tag-version
     (actualiza package.json y package-lock.json a la vez; si se edita
     package.json a mano, el lock queda desfasado y npm lo reescribe
     durante el build, dejándolo pendiente en git)
   - Actualizar a mano la misma versión en (src/app/app.service.ts)
2. Desde powersheel administrador(ir a la ruta del repositorio)
3. cd C:\Users\gonza\OneDrive\Nodus\Proyectos\Nodus-app\parche-gestoria

Powershell administrador

1. ng build
2. npm run build

Una vez esté construído, revisar el .exe en la carpeta;
"C:\Users\gonza\OneDrive\Nodus\Proyectos\Nodus-app\parche-gestoria\dist"

Enviar a través de : https://wetransfer.com/

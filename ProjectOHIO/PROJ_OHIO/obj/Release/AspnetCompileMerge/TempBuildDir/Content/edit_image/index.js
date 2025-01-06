

const fileInput = document.querySelector("#upload");

// enabling drawing on the blank canvas
drawOnImage();

fileInput.addEventListener("change", async (e) => {
    const [file] = fileInput.files;

    // displaying the uploaded image
    const image = document.createElement("img");
    image.src = await fileToDataUri(file);

    // enabling the brush after after the image
    // has been uploaded
    image.addEventListener("load", () => {
        drawOnImage(image);
    });

    return false;
});

function fileToDataUri(field) {
    return new Promise((resolve) => {
        const reader = new FileReader();

        reader.addEventListener("load", () => {
            resolve(reader.result);
        });

        reader.readAsDataURL(field);
    });
}

function setImageFile(_ruta) {
    //alert("ruta " + _ruta);
    // displaying the uploaded image
    const image = document.createElement("img");
    image.src = "";
    //$('#img').setAttribute("src", "");

    image.src = _ruta;

    // enabling the brush after after the image
    // has been uploaded
    image.addEventListener("load", () => {
        //alert(11);
        drawOnImage(image);
    });

}

const sizeElement = document.querySelector("#sizeRange");
let size = sizeElement.value;
sizeElement.oninput = (e) => {
    size = e.target.value;
};

const colorElement = document.getElementsByName("colorRadio");
let color;
colorElement.forEach((c) => {
    if (c.checked) color = c.value;
});

colorElement.forEach((c) => {
    c.onclick = () => {
        color = c.value;
    };
});

const canvasElement2 = document.getElementById("canvas2");
const context2 = canvasElement2.getContext("2d");

function drawOnImage(image = null) {
    const canvasElement = document.getElementById("canvas");
    const context = canvasElement.getContext("2d");

    // if an image is present,
    // the image passed as parameter is drawn in the canvas
    if (image) {
        const imageWidth = image.width;
        const imageHeight = image.height;

        // rescaling the canvas element
        canvasElement.width = imageWidth;
        canvasElement.height = imageHeight;

        context.drawImage(image, 0, 0, imageWidth, imageHeight);


        canvasElement2.width = imageWidth;
        canvasElement2.height = imageHeight;

        context2.drawImage(image, 0, 0, imageWidth, imageHeight);
        
    }

    const clearElement = document.getElementById("clear");
    clearElement.onclick = () => {
        //context.clip();
        context.clearRect(0, 0, canvasElement.width, canvasElement.height);
        //context.restore();
        context.drawImage(canvasElement2, 0, 0)
    };

    let isDrawing;

    canvasElement.onmousedown = (e) => {
        isDrawing = true;
        context.beginPath();
        context.lineWidth = size;
        context.strokeStyle = color;
        context.lineJoin = "round";
        context.lineCap = "round";
        if (screen.width < 500) {
            context.moveTo(e.clientX-9, e.clientY - 55);
        } else {
            context.moveTo(e.clientX - 81, e.clientY - 57 );
        }
        
    };

    canvasElement.onmousemove = (e) => {
        if (isDrawing) {
            if (screen.width < 500) {
                context.lineTo(e.clientX-9, e.clientY - 55);
            } else {
                context.lineTo(e.clientX - 81, e.clientY - 57);
            }
            
            context.stroke();
            context.save();
        }
    };

    canvasElement.onmouseup = function () {
        isDrawing = false;
        context.closePath();
    };

    
}


var limpiar = document.getElementById("limpiar");
limpiar.addEventListener("click", function () {
    const canvasElement = document.getElementById("canvas");
    canvasElement.width = canvasElement.width;
    alert("limpiando");
    //miCanvas.width = miCanvas.width;
}, false);
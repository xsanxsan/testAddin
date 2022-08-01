import React, { useEffect, useState } from "react";
import InsertLocation = Word.InsertLocation;

/* global Word, require */

function App() {
  const [randomImages, setRandomImages] = useState([]);
  const [imageToInsert, setImageToInsert] = useState("");

  useEffect(() => {
    if (imageToInsert) {
      test(imageToInsert);
    }
  }, [imageToInsert]);

  const test = (imageToInsert) => {
    setImageToInsert("");
    return Word.run(async (context) => {
      const selection = context.document.getSelection().load("text");
      //Insert at the end of selection
      selection.insertInlinePictureFromBase64(imageToInsert as string, InsertLocation.end);
      //Insert at the end of whole document
      // context.document.body.insertInlinePictureFromBase64(test as string, InsertLocation.before);
      await context.sync();
    });
  };

  const toDataURL = async (url) => {
    const response = await fetch(url);
    return response.blob();
  };

  const getBase64 = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => resolve(dataUriToBase64(reader.result.toString()));
      reader.onerror = (error) => reject(error);
    });
  };

  const base64ToDataUri = (base64) => {
    return "data:image/png;base64," + base64;
  };

  const dataUriToBase64 = (dataUri) => {
    return dataUri.replace("data:", "").replace(/^.+,/, "");
  };

  const getRandomImages = async () => {
    let files = [];
    for (let i = 0; i < 6; i++) {
      files.push(await toDataURL("https://picsum.photos/75"));
    }
    setRandomImages(files);
  };

  const onClickRandomImage = async () => {
    await getRandomImages();
  };

  const resizeBase64Image = (img, width, height) => {
    // create an off-screen canvas
    console.log("given img", img);
    var canvas = document.createElement("canvas");
    var ctx = canvas.getContext("2d");
    canvas.width = width;
    canvas.height = height;

    var image = new Image();
    image.onload = function () {
      ctx.drawImage(image, 0, 0, width, height);
      setImageToInsert(dataUriToBase64(canvas.toDataURL()));
    };
    image.src = img;
  };

  const getSelection = () => {
    return Word.run(async (context) => {
      // const search = context.document.search("test");
      const search = context.document.search("test").load("items");
      const selection = context.document.getSelection().load("text");
      await context.sync();
      const selectionSearched = selection.search("test").load("items");

      await context.sync();
      selectionSearched.items.forEach((value) => (value.font.highlightColor = "#FFFF00"));
      await context.sync();
      console.log("selection is", selection.text);
      console.log("search is", search.items);
    });
  };

  const onClickImage = async (image) => {
    let test = await getBase64(image);
    resizeBase64Image(base64ToDataUri(test), 50, 50);
  };

  return (
    <div>
      <h3>Hello world</h3>
      <button onClick={onClickRandomImage}>Random images</button>
      <button onClick={getSelection}>Get current selection</button>
      <br />
      {randomImages.map((image, index) => {
        return (
          <img
            style={{ padding: "4px 4px" }}
            key={index}
            src={URL.createObjectURL(image)}
            onClick={() => onClickImage(image)}
          />
        );
      })}
    </div>
  );
}

export default App;

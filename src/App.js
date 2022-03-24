import axios from "axios";
import { useEffect, useState } from "react";
import { BeatLoader } from "react-spinners";
import "./App.css";
import infoLogo from "./assets/info.svg";

function App() {
  const [isLoading, setLoading] = useState(false);
  const [hasDoc, setHasDoc] = useState(true);

  const getFile = () => {
    const query = new URLSearchParams(window.location.search);
    const doc_url = query.get("doc_url");
    if (doc_url !== null) {
      axios
        .post(
          `https://prod-48.westus.logic.azure.com:443/workflows/7ec2dfcb007d41e181f2e89b9cc53645/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=s-a8MVU25ajsqL36ranGZNwwMrWI0pio6nl-H-Us0PM`,
          {
            doc_url: doc_url,
          },
          {
            responseType: "blob",
          }
        )
        .then((res) => {
          const blobContent = new Blob([res.data], {
            type: "application/pdf",
          });
          const wrapper = document.getElementById("wrapper");
          const iframe = document.createElement("iframe");
          
          iframe.name = "printf";
          iframe.src = URL.createObjectURL(blobContent);
          iframe.style.width = "100%";
          iframe.style.height = "100%";

          iframe.onload = () => {
            window.frames["printf"].print();
          };

          wrapper.appendChild(iframe);
          setLoading(false);
        })
        .catch((err) => {
          setHasDoc(false);
          setLoading(false);
        });
    } else {
      setLoading(false);
      setHasDoc(false);
    }
  };

  useEffect(() => {
    getFile();
  }, []);

  return (
    <div id="wrapper">
      {isLoading && (
        <div id="loader">
          <BeatLoader loading={isLoading} size={30} color={"#9013FE"} />
        </div>
      )}

      {!hasDoc && (
        <div id="has-doc">
          <div id="content">
            <img src={infoLogo} alt="info-icon" />
            <h2>
              Please Pass the Relative Path of Document in Sharepoint with the
              doc_url query string.
            </h2>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;

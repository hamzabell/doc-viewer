import axios from "axios";
import { useEffect, useState } from "react";
import { BeatLoader } from "react-spinners";
import "./App.css";
import infoLogo from "./assets/info.svg";
import error from "./assets/errror.svg";

function App() {
  const [isLoading, setLoading] = useState(false);
  const [hasDoc, setHasDoc] = useState(true);
  const [hasError, setHasError] = useState(false);

  const getFormData = () => {
    const formData = new FormData();
    formData.append("client_id", process.env.REACT_APP_client_id);
    formData.append("client_secret", process.env.REACT_APP_client_secret);
    formData.append(
      "resource",
      `${process.env.REACT_APP_principal}/${process.env.REACT_APP_sharepoint_tenant}@${process.env.REACT_APP_tenant_id}`
    );
    formData.append("grant_type", process.env.REACT_APP_grant_type);

    return formData;
  };

  const getFile = (token) => {
    const headers = {
      Accept: "application/json;odata=verbose",
      Authorization: "Bearer " + token,
    };

    const query = new URLSearchParams(window.location.search);
    const doc_url = query.get("doc_url");
    if (doc_url !== null) {
      axios
        .get(
          `https://${process.env.REACT_APP_sharepoint_tenant}/_api/web/GetFileByServerRelativeUrl('${doc_url}')/$value`,
          {
            headers: headers,
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
    const getFileBinary = () => {
      axios
        .post(
          `${process.env.REACT_APP_tenant_id}/tokens/OAuth/2/`,
          getFormData(),
          {
            headers: {
              "Content-Type": "application/x-www-form-urlencoded",
              Cookie: process.env.REACT_APP_cookie,
            },
          }
        )
        .then((res) => {
          setLoading(true);
          getFile(res.data.access_token);
        })
        .catch((err) => {
          setLoading(false);
          setHasDoc(false);
          setHasError(true);
        });
    };

    getFileBinary();
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

      {hasError && (
        <div id="has-error">
          <div id="content">
            <img src={error} alt="info-icon" />
            <h2>
              An Error occurred while trying get the access token, please
              refresh the page!
            </h2>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;

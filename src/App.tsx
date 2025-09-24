import React, { useState } from 'react';

import WineComponent from "./components/wine/WineComponent"
import WhiskyComponent from "./components/whisky/WhiskyComponent"
import EtcComponent from './components/etc/EtcComponent';


function App() {


  const [selectedType, setSelectedType] = useState<string>("wine") //wine, whisky, etc

  return (
    <div>
      <>
      
      <button onClick={(e)=>{
        e.preventDefault();
        if (selectedType !== "wine"){
          setSelectedType("wine")
        }
      }}>
        {selectedType === "wine" && (<strong>Wine</strong>)}
        {selectedType !== "wine" && (<>Wine</>)}
      </button>{" "}
      <button onClick={(e)=>{
        e.preventDefault();
        if (selectedType !== "whisky"){
          setSelectedType("whisky")
        }
      }}>
        {selectedType === "whisky" && (<strong>Whisky</strong>)}
        {selectedType !== "whisky" && (<>Whisky</>)}
      </button>{" "}
      <button onClick={(e)=>{
        e.preventDefault();
        if (selectedType !== "etc"){
          setSelectedType("etc")
        }
      }}>
        {selectedType === "etc" && (<strong>ETC</strong>)}
        {selectedType !== "etc" && (<>ETC</>)}
      </button>{" "}
      </>
      
      <br/>
      
      <div style={{
        }}>
        {selectedType === "wine" && (<WineComponent/>)}
        {selectedType === "whisky" && (<WhiskyComponent/>)}
        {selectedType === "etc" && (<EtcComponent/>)}
      </div>

    </div>
  );
}

export default App;

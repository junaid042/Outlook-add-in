import React, { useState } from "react";
import { Card, Button, Radio, RadioGroup,useId ,Toaster, useToastController, Toast,ToastTitle,ToastTrigger} from "@fluentui/react-components";


const questions = [
    '¿Has pulsado algún enlace?',
    '¿Has notado algún comportamiento anómalo?',
    '¿Has introducido algún código de usuario o contraseña en las pantallas a las que has llegado?',
    'En las últimas horas o días, ¿Has recibido correos electrónicos similares?',
    '¿Has notado algún patrón inusual en los mensajes que recibes?',
    '¿Has accedido a tu correo electrónico desde algún dispositivo o red diferente a los habituales últimamente?',
    '¿Has realizado alguna acción sobre el correo en estos dispositivos?'
];

const Question = () => {
    const [answers, setAnswers] = useState({});
    const [count, setCount] = useState(1);

    const handleRadioChange = (value, question) => {
        setAnswers(prevAnswers => ({
            ...prevAnswers,
            [question]: value
        }));
    };

    const handleSaveClick = () => {
        // Save the selected value for all questions to local storage
        const updatedAnswers = {};
        questions.forEach(question => {
            updatedAnswers[question] = answers[question] || 'no'; 
        });
        localStorage.setItem("questionData", JSON.stringify(updatedAnswers));
        // var url = "https://codename64.atlassian.net/rest/api/3/issue";
        // var username = "mailto:soporte@codename64.com";
        // var password = "Toledo60*";
        // var token = "ATATT3xFfGF07bd8XnTMnbylr8BmHm_1SB9pjgNMkUmHcH-RqUOKY0WSvFQI63zC9fPKOpchXpD4F5gMM4T5Y9ehVzssi_Om6b2ztDdVkgVtPXZgZjfZOek0kHlIm6ymO2WPVaBxJkKUAquukvBbXHkOVaZ3AOB09okf0gRM6MxnaKaPFjpv7Yw=09DF07C6";
        // var projectId = "PC";
        
        // var data = updatedAnswers
        
        // var xhr = new XMLHttpRequest();
        // xhr.open("POST", url, true);
        // xhr.setRequestHeader("Authorization", "Basic " + btoa(username + ":" + password));
        // xhr.setRequestHeader("Content-Type", "application/json");

        // xhr.onreadystatechange = function () {
        //     if (xhr.readyState === 4 && xhr.status === 200) {
        //         var json = JSON.parse(xhr.responseText);
        //         console.log(json);
        //     }
        // };

        // xhr.send(JSON.stringify(data));

    };

    // Toast 
    const toasterId = useId("toaster");
    const { dispatchToast } = useToastController(toasterId);
    const notify = (question) =>{
       setCount((prevCount)=> prevCount + 1)
      dispatchToast(
        <Toast
        style={{borderLeft: '3px solid rgb(8, 127, 231)',position:'fixed',bottom: '10px',left:'20%'}}
        onDismiss={() => dispatchToast(null)}> 
              <ToastTitle>
              Hiciste clic en "sí" {count} veces   
              </ToastTitle>
          </Toast>,
          { intent: "success" }
      );
    }

    const handleNoClick=(question)=>{
        if(answers[question] == 'Sí'){
            setCount((prevCount)=> prevCount -1)
        }
    }
    return (
        <>
            {questions.map((question, index) => (
                <Card key={index} style={{ margin: '2%', borderLeft: '3px solid rgb(8 127 231)' }}>
                    {question}
                    <RadioGroup
                        layout="horizontal"
                        onChange={(e) => handleRadioChange(e.target.value, question)}
                        selectedValue={answers[question] || 'no'} 
                    >
                         <Toaster toasterId={toasterId}   timeout={800} />
                        <Radio label='Sí' value='Sí' onClick={()=>{
                            if (answers[question] !== "Sí") {
                                notify(question)
                            }
                        }} />
                        <Radio label='No' value='no' onClick={()=>handleNoClick(question)} />
                    </RadioGroup>
                </Card>
            ))}
            <div style={{ display: 'flex', justifyContent: 'center', margin: '4%' }}>
                <Button appearance="primary"  onClick={handleSaveClick}>
                Guardar
                </Button>
            </div>
        </>
    );
};

export default Question;


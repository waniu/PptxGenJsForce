import { LightningElement } from 'lwc';
import { loadScript } from 'lightning/platformResourceLoader';
import PPTXGEN from '@salesforce/resourceUrl/pptxGen';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import getUserData from '@salesforce/apex/CTRL_TestPptxGen.getUserData';

export default class SFtoPptxCheck extends LightningElement {
    isLibraryLoaded = false;
    dataTable;
    slides = [];
    pptx;
    queredData;

    connectedCallback() {
        if(!this.isLibraryLoaded) {
            Promise.all([
                loadScript(this, PPTXGEN),
                this.getData()
            ]).then(() => {
                this.isLibraryLoaded = true;
            }).catch((err) => {
                console.error('Error loading pptxGen static resource');
                console.error(JSON.stringify(err, null, 2));
                this.dispatchEvent(
                    new ShowToastEvent({
                        title: 'Error loading pptxGen',
                        message: JSON.stringify(err, null, 2),
                        variant: 'error',
                    }),
                );
            });
        }
    }

    creteFile() {
        this.pptx = new window.PptxGenJS();
        this.prepareTableData();
        this.setTitleSlide();
        this.addSlideWithTabel();
        this.downloadFile();
    }

    getData() {
        getUserData().then(result => this.queredData = result);
    }

    setTitleSlide() {
        this.pptx.title = 'SF to PPTX poc';
        this.pptx.author = "JAWAN";
        this.pptx.subject = "Check PptxGenJS library capabilities";
        this.pptx.company = "Finitec";

        let titleSlide = this.pptx.addSlide();
        titleSlide.background = { 
            color: "#E20074" 
        };
        titleSlide.addText("SF to PPTX POC", {
            x: 0,
            y: 1,
            w: "100%",
            h: "40%",
            align: "center",
            color: "#eff0f1",
            fill: "#0077F7",
            fontSize: 46,
        });
    }

    addSlideWithTabel() {
        let tableSlide = this.pptx.addSlide();
        tableSlide.addTable(this.dataTable, 
            { 
                align: "left",
                border: { 
                    pt: "1", 
                    color: "BBCCDD" 
                } 
            }
        );
    }

    prepareTableData() {
        let dataTable = Object.entries(this.queredData);
        this.dataTable = dataTable;
    }

    downloadFile() {
        this.pptx.write("base64").then((data) => {
            const aHref = 'data:application/vnd.ms-powerpoint;base64,' + data;
            let downloadElement = document.createElement('a');
            downloadElement.href = aHref;
            downloadElement.target = '_self';
            downloadElement.download = 'Demo PptxGen.pptx';
            const downloadDiv = this.querySelector('.download-div');
            document.body.appendChild(downloadElement);
            downloadElement.click();
        }).catch(err => {
            this.dispatchEvent(
                new ShowToastEvent({
                    title: 'Error creating pptx file',
                    message: JSON.stringify(err, null, 2),
                    variant: 'error',
                }),
            );
        })
    }

}
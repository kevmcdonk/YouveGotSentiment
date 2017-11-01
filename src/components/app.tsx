import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Header } from './header';
import * as $ from 'jquery';
import {
    Rating
  } from 'office-ui-fabric-react/lib/Rating';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';

const textAnalyticsKey = '';
const contentModeratorKey = '';

export interface AppProps {
    title: string;
}

export interface AppState {
    currentStatus: string;
    sentimentText: string;
    sentimentValue: number;
    sentimentRating: number;
    averageSentiment: number;
    numberMails: number;
    keyPhrases: string[];
    autoCorrected: string;
    pii: string[];
    dodgyTerms: string[];
    loading: boolean;
    messages: any[];
}

export class App extends React.Component<AppProps, AppState> {
    //private clientApplication: UserAgentApplication;
   
    constructor(props, context) {
        
        super(props, context);
        

        this.state = {
            currentStatus: 'initial',
            sentimentText: 'undefined',
            sentimentValue: 0,
            sentimentRating: 0,
            averageSentiment: 0,
            numberMails: 0,
            keyPhrases: [],
            autoCorrected: '',
            pii: [],
            dodgyTerms: [],
            loading: false,
            messages: []
        };

        this.getCurrentSentimentValue();
        this.setState({
            currentStatus: 'Page loaded'
        });
    }

    componentDidMount() {

        
    }

    /**
 * Generates a GUID string.
 * @returns {String} The generated GUID.
 * @example af8a8416-6e18-a307-bd9c-f2c947bbb3aa
 * @author Slavik Meltser (slavik@meltser.info).
 * @link http://slavik.meltser.info/?p=142
 */
/*
 guid = () => {
    function _p8(s) {
        var p = (Math.random().toString(16)+"000000000").substr(2,8);
        return s ? "-" + p.substr(0,4) + "-" + p.substr(4,4) : p ;
    }
    return _p8() + _p8(true) + _p8(true) + _p8();
}
*/
    getCurrentSentimentValue = () => {
        var self = this;
        Office.context.mailbox.item.loadCustomPropertiesAsync(function(customPropertyResult){
            console.log(customPropertyResult.value);
            var sentiment = customPropertyResult.value.get('MailSentiment');

            if (sentiment){
                self.setState({
                    currentStatus: 'Sentiment retrieved',
                    sentimentValue: sentiment.sentimentValue,
                    sentimentText: sentiment.sentimentText,
                    sentimentRating: sentiment.sentimentRating

                });
            }
            else {
                self.click();
            }

            var keyPhrases = customPropertyResult.value.get('KeyPhrases');
            
            if (keyPhrases){
                self.setState({
                    currentStatus: 'Key phrases retrieved',
                    keyPhrases: keyPhrases
                });
            }

            var cmi = customPropertyResult.value.get('ContentModeratorInfo');
            if (cmi){
                self.setState({
                    currentStatus: 'PII retrieved',
                    pii: cmi.pii,
                    dodgyTerms: cmi.dodgyTerms
                });
            }
        });
    }

    

    click = async () => {
        this.setState({
            currentStatus: 'Checking...'
        }); 

        
        try {
            var self = this;
            //this.getAccessToken();
            
            Office.context.mailbox.item.body.getAsync(
                Office.CoercionType.Text,
                { asyncContext:"This is passed to the callback" },
                function callback(result) {
           
            var body = {
                "documents": [
                    {
                    "language": "en",
                    "id": "1", //TODO: replace with guid()
                    "text": result.value//"how is everyone feeling? Feeling good"
                    }
                ]
            };
        
            

            //Get sentiments
            $.ajax({
                url: "https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/sentiment",
                beforeSend: function(xhrObj){
                    // Request headers
                    xhrObj.setRequestHeader("Content-Type","application/json");
                    xhrObj.setRequestHeader("Ocp-Apim-Subscription-Key",textAnalyticsKey);
                },
                type: "POST",
                // Request body
                data: JSON.stringify(body),
            })
            .done(function(data) {
                var sentimentScore = Math.round(data.documents[0].score*100);
                var sentimentScoreText = 'Neutral';
                if (sentimentScore < 40) {
                    sentimentScoreText = 'Negative';
                }
                if (sentimentScore > 60) {
                    sentimentScoreText = 'Positive';
                }

                self.setState({
                            currentStatus: 'Sentiment retrieved',
                            sentimentValue: sentimentScore,
                            sentimentText: sentimentScoreText,
                            sentimentRating: sentimentScore/20

                        });
                Office.context.mailbox.item.loadCustomPropertiesAsync(function(customPropertiesResult){
                    console.log(customPropertiesResult);
                    customPropertiesResult.value.set("MailSentiment", {
                        sentimentValue: sentimentScore,
                        sentimentText: sentimentScoreText,
                        sentimentRating: sentimentScore/20
                    });
                    customPropertiesResult.value.saveAsync(function(savedResult){
                        console.log(savedResult);
                    });
                });
                
            })
            .fail(function(err) {
                    self.setState({
                            currentStatus: err.responseText
                        });
            });

            //Get key talking points
            $.ajax({
                url: "https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/keyPhrases",
                beforeSend: function(xhrObj){
                    // Request headers
                    xhrObj.setRequestHeader("Content-Type","application/json");
                    xhrObj.setRequestHeader("Ocp-Apim-Subscription-Key",textAnalyticsKey);
                },
                type: "POST",
                // Request body
                data: JSON.stringify(body),
            })
            .done(function(data) {
                var keyPhraseCollection = [];
                data.documents[0].keyPhrases.forEach(function(keyPhrase){
                    keyPhraseCollection.push(keyPhrase);
                });
                

                self.setState({
                    currentStatus: 'Key phrases loaded',
                    keyPhrases: keyPhraseCollection
                });

                Office.context.mailbox.item.loadCustomPropertiesAsync(function(customPropertiesResult){
                    console.log(customPropertiesResult);
                    customPropertiesResult.value.set("KeyPhrases", keyPhraseCollection);
                    customPropertiesResult.value.saveAsync(function(savedResult){
                        console.log(savedResult);
                    });
                });
            })
            .fail(function(err) {
                self.setState({
                    currentStatus: err.responseText
                });
            });

            $.ajax({
                url: "https://westeurope.api.cognitive.microsoft.com/contentmoderator/moderate/v1.0/ProcessText/Screen/?language=eng&autocorrect=true&PII=true",
                beforeSend: function(xhrObj){
                    // Request headers
                    xhrObj.setRequestHeader("Content-Type","text/plain");
                    xhrObj.setRequestHeader("Ocp-Apim-Subscription-Key",contentModeratorKey);
                },
                type: "POST",
                // Request body
                data: result.value.substring(0,1024),
            })
            .done(function(data) {
                var pii = [];
                
                data.PII.Address.forEach(function(piiFound) {
                    pii.push('Address: ' + piiFound.Detected);
                });
                data.PII.Email.forEach(function(piiFound) {
                    pii.push('Email: ' + piiFound.Detected);
                });
                data.PII.IPA.forEach(function(piiFound) {
                    pii.push('IPA: ' + piiFound.Detected);
                });
                data.PII.Phone.forEach(function(piiFound) {
                    pii.push('Phone: ' + piiFound.Detected);
                });
                var terms = [];
                data.Terms.forEach(function(dodgyTerm){
                    terms.push(dodgyTerm.Term);
                });

                self.setState({
                    currentStatus: 'Content moderation retrieved',
                    pii: pii,
                    dodgyTerms: terms
                });
                Office.context.mailbox.item.loadCustomPropertiesAsync(function(customPropertiesResult){
                    console.log(customPropertiesResult);
                    customPropertiesResult.value.set("ContentModeratorInfo", {
                        pii: pii,
                        dodgyTerms: terms
                    });
                    customPropertiesResult.value.saveAsync(function(savedResult){
                        console.log(savedResult);
                    });
                });
            })
            .fail(function(err) {
                //debugger;
                    self.setState({
                            currentStatus: err.responseText
                        });
            });
            
            });
        }
        catch (err) {
            this.setState({
                currentStatus: 'Errored: ' + err.message
            }); 
        }
        
    }

            /*
                <div>
                    <Button className='ms-welcome__action' buttonType={ButtonType.hero} icon='ChevronRight' onClick={this.clickLoadMessages} >Load messages</Button>
                </div>
                <div>
                    <Button className='ms-welcome__action' buttonType={ButtonType.hero} icon='ChevronRight' onClick={this.getCurrentSentimentValue} >Reload local</Button>
                </div>
                <div>
                    <Button className='ms-welcome__action' buttonType={ButtonType.hero} icon='ChevronRight' onClick={this.clickAverageSentiment} >Load average sentiment</Button>
                </div>
                */
                
    
    render() {
        //Adal.processAdalCallback();
        return (
            <div className='ms-welcome'>
                <Header logo='assets/logo-filled.png' title={this.props.title} message='Welcome'>
                </Header>
                <Pivot>
                    <PivotItem linkText='Sentiment'>
                        <div className="ms-bgColor-neutralLighterAlt">
                            <Label className='ms-font-xl'>How positive is your mail?</Label>
                            <Rating min={ 1 } max={ 5 } rating={ this.state.sentimentRating }  />
                            <Label>Value: {this.state.sentimentRating}%</Label>
                        </div>
                    </PivotItem>
                    <PivotItem linkText='Key phrases'>
                    <ul>
                        {this.state.keyPhrases.map(function(keyPhrase, i){
                            return <li key={i}>{keyPhrase}</li>;
                        })}
                        </ul>
                        
                    </PivotItem>
                    <PivotItem linkText='Content Moderator'>
                    
                        <ul>
                        {this.state.pii.map(function(piiInfo, i){
                            return <li key={i}>{piiInfo}</li>;
                        })}
                        {this.state.dodgyTerms.map(function(term, i){
                            return <li key={i}>Dodgy: {term}</li>;
                        })}
                        </ul>
                    </PivotItem>
                </Pivot>
                
                <div>
                    <Button className='ms-welcome__action' buttonType={ButtonType.hero} icon='ChevronRight' onClick={this.click} >Reload</Button>
                </div>
                
                <div className='ms-font-l'>Current status: <span>{this.state.currentStatus}</span></div>
            </div>
        );
        
    };

    
    
};

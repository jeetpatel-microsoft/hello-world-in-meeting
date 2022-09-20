import React from "react";
import { app } from "@microsoft/teams-js";
import MediaQuery from 'react-responsive';
import './App.css';

class Tab extends React.Component {
  constructor(props){
    super(props)
    this.state = {
      context: {},
      changing_var : "unchanged",
      my_meeting_id : "",
    }
  }

  //React lifecycle method that gets called once a component has finished mounting
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidmount
  componentDidMount(){
    setTimeout(() =>{
      this.setState({changing_var: "changed now"})
    }, 60000);
    app.initialize().then(() => {
      // Get the user context from Teams and set it in the state
      app.getContext().then(async (context) => {
        if(Object.keys(context).length != 0){
          this.setState({my_meeting_id : context.meeting.id})
        }
        this.setState({
          context: context
        });
      });
    });
    // Next steps: Error handling using the error object
  }

  render() {
    let meetingId = this.state.context['meetingId'] ?? "";
    let myMeetingIDContext = null;
    try{
      let myMeetingIDContext =  this.state.my_meeting_id;
    }
    catch(err){
      let myMeetingIDContext = ""
    }
    let userPrincipleName = this.state.context['userPrincipalName'] ?? "";
    console.log(meetingId)
    return (
    <div>
      <h1>In-meeting app sample</h1>
      <h1>{this.state.changing_var}</h1>
      <h3>Principle Name:</h3>
      <p>{userPrincipleName}</p>
      <h3>Meeting ID:{myMeetingIDContext}</h3>

      <MediaQuery maxWidth={280}>
        <h3>This is the side panel</h3>
        <a href="https://docs.microsoft.com/en-us/microsoftteams/platform/apps-in-teams-meetings/teams-apps-in-meetings">Need more info, open this document in new tab or window.</a>
      </MediaQuery>
    </div>
    );
  }
}

export default Tab;

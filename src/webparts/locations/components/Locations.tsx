import * as React from 'react';
import {useState, useEffect} from 'react';
import styles from './Locations.module.scss';
import type { ILocationsProps } from './ILocationsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';

const setID = 'f6fc9802-3af3-4200-92de-9fe5f4af4fd2'; // offices term set
const termID = 'e8e6feb5-1cf7-47bd-afb7-e352b78bd365'; // division terms

const Locations : React.FunctionComponent<ILocationsProps> = (props) => {
  const [terms,setTerms] = useState([]);
  const [selectedTerm, setSelectedTerm] = useState('');

  useEffect(() => {
    const fetchTerms=async() => {
      const url : string = `${props.context.pageContext.web.absoluteUrl}/_api/v2.1/termStore/groups/a66b7b2f-9f5d-4573-b763-542518574351/sets/${setID}/terms/${termID}/children?select=*`;
      const response = await props.context.spHttpClient.get(url,SPHttpClient.configurations.v1);
      if(response.ok){
        const data = await response.json();
        setTerms(data.value);        
      } else {
        console.error('Error fetching terms:', response.statusText);
      }
    };
    fetchTerms();

  },[props.context,setID,termID]);

  const handleTermChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
    setSelectedTerm(event.target.value);
  };

/*
  const renderTerms = (items:any) => {
    console.log("terms",items);
    return (
      <ul className={styles.links}>  
        {items.map((term: any) => (
          <li key={term.id}>{term.labels[0].name}</li>
        ))}
      </ul>
    )
  }
*/

  const {
    description,
    isDarkTheme,
    environmentMessage,
    hasTeamsContext,
    userDisplayName
  } = props;

  return (
    <section className={`${styles.locations} ${hasTeamsContext ? styles.teams : ''}`}>
    <div className={styles.welcome}>
      <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
      <h2>Well done, {escape(userDisplayName)}!</h2>
      <div>{environmentMessage}</div>
      <div>Web part property value: <strong>{escape(description)}</strong></div>
    </div>
    <div>
      <h3>Welcome to SharePoint Framework!</h3>
      <p>
        The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
      </p>
      <div>
        <h2>Term Store Terms</h2>
          {terms.length > 0 ? (
            <div>
              <select onChange={handleTermChange} value={selectedTerm}>
                <option value="">Select a term</option>
                {terms.map((term: any) => (
                  <option key={term.id} value={term.id}>{term.labels[0].name}</option>
                ))}
              </select>
              {selectedTerm && (
                <div>
                  <h3>Selected Term Details</h3>
                  {/*<p>{terms.find(term => term.id === selectedTerm)?.labels[0].name}</p>*/}
                </div>
              )}
            </div>
          ) : (
            <p>Loading terms...</p>
          )}
      </div>
    </div>
  </section>

  )
}

export default Locations;
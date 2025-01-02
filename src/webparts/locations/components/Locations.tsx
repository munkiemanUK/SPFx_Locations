import * as React from 'react';
import {useState, useEffect} from 'react';
import styles from './Locations.module.scss';
import type { ILocationsProps } from './ILocationsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';

const setID = 'f6fc9802-3af3-4200-92de-9fe5f4af4fd2'; // offices term set
//const termID = 'e8e6feb5-1cf7-47bd-afb7-e352b78bd365'; // division terms

interface Term {
  id: string;
  labels: { name: string }[];
}

const Locations: React.FunctionComponent<ILocationsProps> = (props) => {
  const [level1Terms, setLevel1Terms] = useState<Term[]>([]);
  const [level2Terms, setLevel2Terms] = useState<Term[]>([]);
  const [level3Terms, setLevel3Terms] = useState<Term[]>([]);
  const [selectedLevel1Term, setSelectedLevel1Term] = useState<string>('e8e6feb5-1cf7-47bd-afb7-e352b78bd365'); // initial division terms
  const [selectedLevel2Term, setSelectedLevel2Term] = useState<string>('');
  const [selectedLevel3Term, setSelectedLevel3Term] = useState<string>('');

  const fetchTerms = async (termID: string, setTerms: React.Dispatch<React.SetStateAction<Term[]>>) => {
    const url: string = `${props.context.pageContext.web.absoluteUrl}/_api/v2.1/termStore/groups/a66b7b2f-9f5d-4573-b763-542518574351/sets/${setID}/terms/${termID}/children?select=*`;
    const response = await props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (response.ok) {
      const data = await response.json();
      setTerms(data.value);
    } else {
      console.error('Error fetching terms:', response.statusText);
    }
  };

  useEffect(() => {
    fetchTerms(selectedLevel1Term, setLevel1Terms);
  }, [props.context, selectedLevel1Term]);

  useEffect(() => {
    if (selectedLevel1Term) {
      fetchTerms(selectedLevel1Term, setLevel2Terms);
    }
  }, [selectedLevel1Term]);

  useEffect(() => {
    if (selectedLevel2Term) {
      fetchTerms(selectedLevel2Term, setLevel3Terms);
    }
  }, [selectedLevel2Term]);

  const handleLevel1Change = (event: React.ChangeEvent<HTMLSelectElement>) => {
    const newTermID = event.target.value;
    setSelectedLevel1Term(newTermID);
    setSelectedLevel2Term('');
    setSelectedLevel3Term('');
    setLevel2Terms([]);
    setLevel3Terms([]);
  };

  const handleLevel2Change = (event: React.ChangeEvent<HTMLSelectElement>) => {
    const newTermID = event.target.value;
    setSelectedLevel2Term(newTermID);
    setSelectedLevel3Term('');
    setLevel3Terms([]);
  };

  const handleLevel3Change = (event: React.ChangeEvent<HTMLSelectElement>) => {
    const newTermID = event.target.value;
    setSelectedLevel3Term(newTermID);
  };

  const renderTerms = (items: Term[]) => {
    return (
      <ul className={styles.links}>
        {items.map((term: Term) => (
          <li key={term.id}>{term.labels[0].name}</li>
        ))}
      </ul>
    );
  };

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
          <div>
            <select onChange={handleLevel1Change} value={selectedLevel1Term}>
              <option value="">Select a term</option>
              {level1Terms.map((term: Term) => (
                <option key={term.id} value={term.id}>{term.labels[0].name}</option>
              ))}
            </select>
            {selectedLevel1Term && (
              <select onChange={handleLevel2Change} value={selectedLevel2Term}>
                <option value="">Select a term</option>
                {level2Terms.map((term: Term) => (
                  <option key={term.id} value={term.id}>{term.labels[0].name}</option>
                ))}
              </select>
            )}
            {selectedLevel2Term && (
              <select onChange={handleLevel3Change} value={selectedLevel3Term}>
                <option value="">Select a term</option>
                {level3Terms.map((term: Term) => (
                  <option key={term.id} value={term.id}>{term.labels[0].name}</option>
                ))}
              </select>
            )}
          </div>
          {selectedLevel3Term && (
            <div>
              <h3>Selected Term Details</h3>
              <p>{level3Terms.find(term => term.id === selectedLevel3Term)?.labels[0].name}</p>
            </div>
          )}
          {renderTerms(level3Terms)}
        </div>
      </div>
    </section>
  );
};

export default Locations;
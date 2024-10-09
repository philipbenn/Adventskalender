import * as React from 'react';
import { Card } from '@fluentui/react-components';
import { IAdventsKalenderProps } from './IAdventsKalenderProps';
import styles from './AdventsKalender.module.scss';

const AdventsKalender = (props: IAdventsKalenderProps): JSX.Element => {
  const currentDate = new Date().getTime();
  const currentYear = new Date().getFullYear();

  // Utility functions
  const parseDate = (dateString: string): number => new Date(dateString).getTime();
  const isClicked = (adventKey: string): boolean => localStorage.getItem(`${adventKey}${currentYear}`) === 'true';
  const markAsClicked = (adventKey: string): void => localStorage.setItem(`${adventKey}${currentYear}`, 'true');

  const isClickable = (adventDate: number, adventKey: string): boolean => {
    return currentDate >= adventDate && !isClicked(adventKey);
  };

  const handleClick = (adventUrl: string, adventDate: number, adventKey: string): void => {
    if (isClickable(adventDate, adventKey)) {
      markAsClicked(adventKey);
      window.location.href = adventUrl;
    }
  };

  // Data for advents
  const advents = [
    { key: 'førsteAdvent', label: props.førsteAdventTitle, date: parseDate(props.førsteAdventDateTime.value.toString()), url: props.førsteAdventUrl },
    { key: 'andenAdvent', label: props.andenAdventTitle, date: parseDate(props.andenAdventDateTime.value.toString()), url: props.andenAdventUrl },
    { key: 'tredjeAdvent', label: props.tredjeAdventTitle, date: parseDate(props.tredjeAdventDateTime.value.toString()), url: props.tredjeAdventUrl },
    { key: 'fjerdeAdvent', label: props.fjerdeAdventTitle, date: parseDate(props.fjerdeAdventDateTime.value.toString()), url: props.fjerdeAdventUrl },
  ];

  return (
    <>
    <h1 className={styles.title}>{props.title}</h1>
    <div 
      className={styles.adventCalenderContainer} 
      style={{ backgroundImage: `url(${props.backgroundImageUrl})`, backgroundSize: 'cover', backgroundPosition: 'center' }}
    >
      {advents.map((advent) => (
        <Card
          key={advent.key}
          className={`${styles.adventBox} ${isClickable(advent.date, advent.key) ? styles.clickable : isClicked(advent.key) ? styles.clicked : styles.disabled}`}
          onClick={() => handleClick(advent.url, advent.date, advent.key)}
        >
          <h2>{advent.label}</h2>
        </Card>
      ))}
    </div>
    </>
  );
};

export default AdventsKalender;
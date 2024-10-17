import * as React from 'react';
import { Card } from '@fluentui/react-components';
import { IconButton } from '@fluentui/react/lib/Button';
import { IAdventsCalendarProps } from '../interface/IAdventsKalenderProps';
import { Modal, Stack } from '@fluentui/react';
import { useBoolean } from '@fluentui/react-hooks';
import AdventsService from '../service/AdventsService';
import styles from './AdventsKalender.module.scss';

const AdventsCalendar = (props: IAdventsCalendarProps): JSX.Element => {
  
  // Constants
  const currentDate = new Date().getTime();
  const textColorStyle = { color: props.textColor };

  // Hooks
  const [isModalOpen, { setTrue: showModal, setFalse: hideModal }] = useBoolean(false);
  const [selectedAdventUrl, setSelectedAdventUrl] = React.useState<string | null>(null);
  const [showCalender , setShowCalender] = React.useState(false);

  const init = async (): Promise<void> => {
    const groupMember = await AdventsService.getGroupMember(props.context, props.group[0].fullName, props.userDisplayName);
    setShowCalender(!!groupMember.value);
  };

  React.useEffect(() => {
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    init();
  }, []);

  // Utility functions
  const parseDate = (dateString: string): number => new Date(dateString).getTime();

  const isClickable = (adventDate: number): boolean => {
    return currentDate >= adventDate;
  };

  // Data for advents
  const items = [
    {
      key: 'firstItem',
      label: props.firstItemTitle,
      date: parseDate(props.firstItemDateTime?.value.toString()),
      url: props.firstItemUrl,
      imageUrl: props.firstItemImage?.fileAbsoluteUrl,
    },
    {
      key: 'secondItem',
      label: props.secondItemTitle,
      date: parseDate(props.secondItemDateTime?.value.toString()),
      url: props.secondItemUrl,
      imageUrl: props.secondItemImage?.fileAbsoluteUrl,
    },
    {
      key: 'thirdItem',
      label: props.thirdItemTitle,
      date: parseDate(props.thirdItemDateTime?.value.toString()),
      url: props.thirdItemUrl,
      imageUrl: props.thirdItemImage?.fileAbsoluteUrl,
    },
    {
      key: 'fourthItem',
      label: props.fourthItemTitle,
      date: parseDate(props.fourthItemDateTime?.value.toString()),
      url: props.fourthItemUrl,
      imageUrl: props.fourthItemImage?.fileAbsoluteUrl,
    },
  ];

  const getClosestItem = (): typeof items[0] | undefined => {
    const pastAdvents = items.filter((item) => item.date <= currentDate);

    if (pastAdvents.length === 0) {
      return undefined;
    }

    const closestItem = pastAdvents.sort((a, b) => currentDate - a.date - (currentDate - b.date))[0];

    return closestItem;
  };

  const closestAdvent = getClosestItem();

  const handleCardClick = (adventUrl: string): void => {
    setSelectedAdventUrl(adventUrl);
    showModal();
  };

  return (
    <>
      {showCalender && (
        <div className={styles.outerContainer}>
          <h2 className={styles.title}>{props.title}</h2>
          <div
            className={styles.adventCalenderContainer}
            style={{
              backgroundImage: `url(${props.backgroundImageUrl})`,
              backgroundSize: 'cover',
              backgroundPosition: 'center',
            }}
          >
            {props.isOneElement && closestAdvent && (
              <Card
                key={closestAdvent.key}
                className={`${styles.adventBox} ${
                  isClickable(closestAdvent.date) ? styles.clickable : styles.disabled
                }`}
                onClick={() => handleCardClick(closestAdvent.url)}
                style={{
                  backgroundImage: `url(${closestAdvent.imageUrl})`,
                  backgroundSize: 'cover',
                  backgroundPosition: 'center',
                }}
              >
                <h2 style={textColorStyle}>{closestAdvent.label}</h2>
              </Card>
            )}

            {!props.isOneElement &&
              items.map((item) => (
                <Card
                  key={item.key}
                  className={`${styles.adventBox} ${
                    isClickable(item.date) ? styles.clickable : styles.disabled
                  }`}
                  style={{
                    backgroundImage: `url(${item.imageUrl})`,
                    backgroundSize: 'cover',
                    backgroundPosition: 'center',
                  }}
                  onClick={() => handleCardClick(item.url)}
                >
                  <h2 style={textColorStyle}>{item.label}</h2>
                </Card>
              ))}
          </div>
        </div>
      )}

      <Modal isOpen={isModalOpen} onDismiss={hideModal}>
        <div className={styles.modalHeader}>
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            ariaLabel="Close modal"
            onClick={hideModal}
            className={styles.closeButton}
          />
        </div>
        <Stack>
          <iframe
            width="100%"
            height="700px"
            src={selectedAdventUrl || ''}
            style={{ border: 'none', maxWidth: '100%', maxHeight: '100vh' }}
            allowFullScreen
          />
        </Stack>
      </Modal>
    </>
  );
};

export default AdventsCalendar;
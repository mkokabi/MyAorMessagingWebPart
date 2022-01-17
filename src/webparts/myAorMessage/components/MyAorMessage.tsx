import * as React from "react";
import styles from "./MyAorMessage.module.scss";
import { IMyAorMessageProps } from "./IMyAorMessageProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";
import { chain } from "lodash";
import { useEffect, useState } from "react";

export type TeamChannel = {
  key: string;
  value: string;
};

export const MyAorMessage = (props: IMyAorMessageProps) => {
  const [teamChannels, setTeamChannels] = useState([]);

  useEffect(() => {
    const _teamChannels: TeamChannel[] = [];

    props.messagingClient
      .get(
        "https://aor-MYAORMSG2-dev01-func.azurewebsites.net/api/channels",
        AadHttpClient.configurations.v1
      )
      .then((res: HttpClientResponse): Promise<any> => {
        return res.json();
      })
      .then((response: TeamChannel[]): void => {
        response.forEach((ch) => {
          _teamChannels.push({ key: ch.key, value: ch.value });
          console.log(`key: ${ch.key} value:${ch.value}`);
        });
        setTeamChannels(_teamChannels);
      });
  }, []);

  return (
    <div className={styles.myAorMessage}>
      <div className={styles.container}>
        <div className={styles.row}>
          <div className={styles.column}>
            <span className={styles.title}>Welcome to SharePoint!</span>
            <p className={styles.subTitle}>
              Customize SharePoint experiences using Web Parts.
            </p>
            <p className={styles.description}>
              {escape(props.description)}
            </p>
            <a href="https://aka.ms/spfx" className={styles.button}>
              <span className={styles.label}>Learn more</span>
            </a>
            {teamChannels.map((ch) => (
              <p>
                {ch.key} {ch.value}
              </p>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

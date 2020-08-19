import * as React from "react";
import getStore from "../../store/SummaryStore";
import { Flex, Loader, FocusZone, ListItem, Avatar } from '@fluentui/react-northstar';
import { observer } from "mobx-react";
import { fetchNonReponders } from "../../actions/SummaryActions";
import { ProgressState } from "./../../utils/SharedEnum";
import { RecyclerViewComponent, RecyclerViewType } from "../RecyclerViewComponent";
import { UxUtils } from "./../../utils/UxUtils";

interface IUserInfoViewProps {
    userName: string;
    accessibilityLabel?: string;
}

/**
 * <NonResponderView> component for the non-responders tab
 */
@observer
export class NonResponderView extends React.Component {
    componentWillMount() {
        fetchNonReponders();
    }

    render() {
        let rowsWithUser: IUserInfoViewProps[] = [];
        if (getStore().progressStatus.nonResponder == ProgressState.InProgress) {
            return <Loader />;
        }
        if (getStore().progressStatus.nonResponder == ProgressState.Completed) {
            for (let userProfile of getStore().nonResponders) {
                userProfile = getStore().userProfile[userProfile.id];

                if (userProfile) {
                    rowsWithUser.push({
                        userName: userProfile.displayName,
                        accessibilityLabel: userProfile.displayName,
                    });
                }
            }
        }
        return (
            <FocusZone className="zero-padding" isCircularNavigation={true}>
                <Flex column className="list-container" gap="gap.small">
                    <RecyclerViewComponent
                        data={rowsWithUser}
                        rowHeight={48}
                        onRowRender={(
                            type: RecyclerViewType,
                            index: number,
                            userProps: IUserInfoViewProps
                        ): JSX.Element => {
                            return (
                                <div aria-label={userProps.accessibilityLabel}  {...UxUtils.getListItemProps()}>
                                    <ListItem className="zero-padding"
                                        index={index}
                                        media={<Avatar name={(userProps.userName).toUpperCase()} size="medium" aria-hidden="true" />}
                                        header={userProps.userName}
                                    />
                                </div>
                            )
                        }}
                    />
                </Flex>
            </FocusZone>
        );
    }
}

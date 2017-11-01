import * as React from 'react';

export interface HeaderProps {
    title: string;
    logo: string;
    message: string;
}

export class Header extends React.Component<HeaderProps, any> {
    constructor(props, context) {
        super(props, context);
    }

    render() {
        return (
            <section className='ms-welcome__header ms-bgColor-yellow ms-u-fadeIn500'>
                <div className="ms-Grid">
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-u-sm6 ms-u-md6 ms-u-lg6 iconAlignRight"><i className="ms-Icon ms-Icon--Emoji ms-fontSize-su"></i></div>
                        <div className="ms-Grid-col ms-u-sm6 ms-u-md6 ms-u-lg6"><i className="ms-Icon ms-Icon--OutlookLogo ms-fontSize-su"></i></div>
                    </div>
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"><span className='ms-fontSize-xxl ms-fontWeight-light ms-fontColor-neutralPrimary'>{this.props.title}</span></div>
                        </div>
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-u-sm6 ms-u-md6 ms-u-lg6 iconAlignRight"><i className="ms-Icon ms-Icon--ChatInviteFriend ms-fontSize-su"></i></div>
                        <div className="ms-Grid-col ms-u-sm6 ms-u-md6 ms-u-lg6"><i className="ms-Icon ms-Icon--Like ms-fontSize-su"></i></div>
                    </div>
                </div>
                <h2 className='ms-fontSize-xl ms-fontWeight-light ms-fontColor-neutralPrimary'>{this.props.message}</h2>                
            </section>
        );
    };
};

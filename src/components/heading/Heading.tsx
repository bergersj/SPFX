import * as React from 'react';
import "./heading.module.scss";


export interface ILink {
    text: string;
    url: string;
}

export interface IHeadingProps {
    heading: string;
    leftLink?: ILink;
    rightLink?: ILink;
    hasBackgroundColor?: boolean;
}

export default class Heading extends React.Component<IHeadingProps> {
    public render() {
        const { heading,leftLink, rightLink, hasBackgroundColor } = this.props;
        const bgColor = hasBackgroundColor ? "bg" : "";
        return (
            <div className={`sectionHeading ${bgColor}`}>
                <div className="heading">{heading}</div>
                {
                    leftLink && leftLink.text &&
                    <div className="leftLink">
                        <a href={leftLink.url || "#"}>{leftLink.text}</a>
                    </div>
                }
                {
                    rightLink && rightLink.text &&
                    <div className="rightLink">
                        <a href={rightLink.url || "#"}>{rightLink.text}</a>
                    </div>
                }
            </div>
        );
    }
}
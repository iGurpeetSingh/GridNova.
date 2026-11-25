import type { CSSProperties } from "react";
import type { FloatingElementData } from "../types";
const paths:Record<string,string>={
 rectangle:"M4 5H25V24H4Z",ellipse:"M3 14A11 10 0 1 0 25 14A11 10 0 1 0 3 14",
 line:"M4 24L26 6",arrow:"M4 24L24 6M17 6H24V13","line-arrow":"M4 24L24 6M17 6H24V13","double-arrow":"M5 24L24 6M5 24V17M5 24H12M24 6V13M24 6H17",
 "elbow-connector":"M5 5V22H24M5 5H13","curved-line":"M3 24C8 2 18 2 25 22",
 "rounded-rectangle":"M8 4H21Q25 4 25 8V21Q25 25 21 25H8Q4 25 4 21V8Q4 4 8 4Z",
 "snip-single-corner":"M4 10L10 4H25V25H4Z","triangle":"M14 3L26 25H2Z","right-triangle":"M4 4V25H25Z",
 diamond:"M14 2L26 14L14 26L2 14Z",pentagon:"M14 2L26 10L22 25H6L2 10Z",hexagon:"M8 3H20L27 14L20 25H8L1 14Z",
 star:"M14 2L17 10L26 10L19 16L22 25L14 20L6 25L9 16L2 10L11 10Z",
 heart:"M14 25C10 21 3 16 3 10A6 6 0 0 1 14 7A6 6 0 0 1 25 10C25 16 18 21 14 25Z",
 "plus-cross":"M11 3H17V11H25V17H17V25H11V17H3V11H11Z",
 "right-arrow":"M2 9H16V3L27 14L16 25V19H2Z","left-arrow":"M26 9H12V3L1 14L12 25V19H26Z",
 "up-arrow":"M9 26V12H3L14 1L25 12H19V26Z","down-arrow":"M9 2V16H3L14 27L25 16H19V2Z",
 "left-right-arrow":"M8 8H20V3L27 14L20 25V20H8V25L1 14L8 3Z",
 "rect-callout":"M3 4H26V19H15L9 26V19H3Z","rounded-rect-callout":"M8 4H21Q26 4 26 9V18Q26 22 21 22H14L8 27V22Q3 22 3 18V9Q3 4 8 4Z",
 "oval-callout":"M14 3C7 3 2 7 2 13S7 22 14 22H17L23 27V20C26 17 27 14 26 11C25 6 20 3 14 3Z",
 "cloud-callout":"M7 20Q1 19 2 13Q2 7 8 7Q10 1 17 4Q22 1 25 7Q29 8 26 14Q27 20 20 20H12L7 26Z",
};
const names:Record<string,string>={rectangle:"Rectangle",ellipse:"Oval",line:"Line",arrow:"Arrow","line-arrow":"Line Arrow","double-arrow":"Double Arrow","elbow-connector":"Elbow Connector","curved-line":"Curved Line","rounded-rectangle":"Rounded Rectangle","snip-single-corner":"Snip Single Corner",triangle:"Triangle","right-triangle":"Right Triangle",diamond:"Diamond",pentagon:"Pentagon",hexagon:"Hexagon",star:"Star",heart:"Heart","plus-cross":"Plus / Cross","right-arrow":"Right Arrow","left-arrow":"Left Arrow","up-arrow":"Up Arrow","down-arrow":"Down Arrow","left-right-arrow":"Left-Right Arrow","rect-callout":"Rectangular Callout","rounded-rect-callout":"Rounded Rectangular Callout","oval-callout":"Oval Callout","cloud-callout":"Cloud Callout"};
const lines=new Set(["line","arrow","line-arrow","double-arrow","elbow-connector","curved-line"]);
export const officeShapeName=(shape?:FloatingElementData["shape"])=>shape?names[shape]??"Shape":"Rectangle";
export default function OfficeShape({shape="rectangle",fill="#DCEEFF",stroke="#0F6CBD",style}:{shape?:FloatingElementData["shape"];fill?:string;stroke?:string;style?:CSSProperties}){
 return <svg viewBox="0 0 28 28" preserveAspectRatio="none" className="h-full w-full p-2" style={style} aria-hidden="true"><path d={paths[shape]??paths.rectangle} fill={lines.has(shape)?"none":fill} stroke={stroke} strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/></svg>;
}

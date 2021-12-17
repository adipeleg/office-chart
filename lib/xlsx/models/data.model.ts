export interface IData {
    type: 'line' | 'bar';
    title: {
        name: string,
        color?: string,
        size?: number
    };
    range: string,
    rgbColors?: string[],
    labels?: boolean,
    marker?: {
        size?: number;
        shape?: string;
    }
}
export interface IData {
    type: 'line' | 'bar';
    title: {
        name: string;
        color?: string;
    };
    range: string;
    rgbColors?: string[];
    marker?: {
        size?: number;
        shape?: string;
    };
}

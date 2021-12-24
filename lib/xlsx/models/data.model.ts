export interface IData {
    type: 'line' | 'bar' | 'pie' | 'scatter';
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

export interface IPPTChartData {
    type: 'line' | 'bar';
    title: {
        name: string,
        color?: string,
        size?: number
    };
    range?: string,
    data?: any[][],
    rgbColors?: string[],
    labels?: boolean,
    marker?: {
        size?: number;
        shape?: string;
    }
}
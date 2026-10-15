import type {Operation} from './analytics';
export type Point = {date: string | null; total: number | null; fixedUsdAmount: number | null; fixedIssue: string | null; row: number; issue: string | null};
export type Data = {sheet: string; name: string; updated: string; operations: Operation[]; points: Point[]};

module Types

type WellInfo =
    { Well: string
      Pad: string
      Process: string 
      DateStart : string
      Duration : string}


type Well(info: WellInfo) =
    member x.Info = info



module WellInfoConstructor =
    let Create well pad state datestart duration =
        { Well = string well
          Pad = string pad
          Process = string state 
          DateStart = string datestart
          Duration =string duration }